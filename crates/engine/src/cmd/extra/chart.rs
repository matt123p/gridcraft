use super::{
    CellRef, Chart, Color, EngineError, Json, LegendPos, RangeRef, Result, Session, Trendline, TrendlineKind, bad, bool_param, cell_param_on, edit,
    f64_param, json, split_sheet, str_param, target_sheet, u32_param,
};

fn find_chart(s: &Session, p: &Json) -> Result<(usize, usize)> {
    let d = s.doc()?;
    let si = if p.get("sheet").is_some() { target_sheet(s, p)? } else { d.wb.active_sheet() };
    let sh = d.wb.sheet(si).ok_or(EngineError::NoDocument)?;
    let idx = match u32_param(p, "chart") {
        Some(id) => sh.charts.iter().position(|c| c.id == id),
        None => sh.charts.len().checked_sub(1),
    };
    idx.map(|i| (si, i)).ok_or_else(|| bad("chart", "no chart on this sheet"))
}

fn with_chart(s: &mut Session, p: &Json, f: impl FnOnce(&mut Chart) -> Result<()>) -> Result<Json> {
    let (si, ci) = find_chart(s, p)?;
    edit(s, |cx| {
        let c = cx.sheet_mut(si)?.charts.get_mut(ci).ok_or_else(|| bad("chart", "no such chart"))?;
        f(c)?;
        Ok(json!({"chart": c.id}))
    })
}

pub(super) fn add_element(s: &mut Session, p: &Json) -> Result<Json> {
    let el = str_param(p, "element").unwrap_or("title").to_string();
    let on = bool_param(p, "on").unwrap_or(true);
    let pos = str_param(p, "position").map(str::to_string);
    with_chart(s, p, |c| {
        match el.as_str() {
            "title" => c.title = if on { c.title.clone().filter(|t| !t.is_empty()).or(Some("Chart Title".into())) } else { Some(String::new()) },
            "legend" => {
                c.legend = if !on {
                    LegendPos::None
                } else {
                    match pos.as_deref() {
                        Some("top") => LegendPos::Top,
                        Some("left") => LegendPos::Left,
                        Some("right") => LegendPos::Right,
                        _ => LegendPos::Bottom,
                    }
                }
            }
            "dataLabels" => c.data_labels = on,
            "gridlines" => c.gridlines = on,
            "axisTitles" => {
                c.x_title = on.then(|| c.x_title.clone().unwrap_or_else(|| "Axis Title".into()));
                c.y_title = on.then(|| c.y_title.clone().unwrap_or_else(|| "Axis Title".into()));
            }
            "smooth" => {
                for s in &mut c.series {
                    s.smooth = on;
                }
            }
            "trendline" => {
                let only = u32_param(p, "series").map(|i| i as usize);
                if only.is_some_and(|i| i >= c.series.len()) {
                    return Err(bad("chart.addElement", "no such series"));
                }
                let targets: Vec<usize> = (0..c.series.len()).filter(|i| only.is_none_or(|o| o == *i)).collect();
                if !on {
                    for i in targets {
                        if let Some(se) = c.series.get_mut(i) {
                            if let Some(index) = u32_param(p, "trendline") {
                                if (index as usize) >= se.trendlines.len() {
                                    return Err(bad("chart.addElement", "no such trendline"));
                                }
                                se.trendlines.remove(index as usize);
                            } else {
                                se.trendlines.clear();
                            }
                        }
                    }
                    return Ok(());
                }
                if let Some(index) = u32_param(p, "trendline") {
                    let i = only.ok_or_else(|| bad("chart.addElement", "trendline edit requires series"))?;
                    let old = c
                        .series
                        .get(i)
                        .and_then(|se| se.trendlines.get(index as usize))
                        .ok_or_else(|| bad("chart.addElement", "no such trendline"))?
                        .clone();
                    let mut t = old;
                    if p.get("type").is_some() {
                        t.kind = trendline_param(p)?.kind;
                    }
                    if let Some(v) = f64_param(p, "forward") {
                        t.forward = Trendline::clamp_forecast(v);
                    }
                    if let Some(v) = f64_param(p, "backward") {
                        t.backward = Trendline::clamp_forecast(v);
                    }
                    if let Some(v) = p.get("intercept") {
                        t.intercept = v.as_f64();
                    }
                    if let Some(v) = p.get("name") {
                        t.name = v.as_str().map(str::to_owned);
                    }
                    if let Some(value) = p.get("color") {
                        t.color = Some(serde_json::from_value(value.clone()).map_err(|_| bad("chart.addElement", "invalid native color"))?);
                    }
                    if let Some(value) = p.get("width") {
                        let width =
                            value.as_f64().filter(|n| (0.25..=12.0).contains(n)).ok_or_else(|| bad("chart.addElement", "invalid line width"))?;
                        t.width = Some(width as f32);
                    }
                    if let Some(value) = p.get("dash") {
                        t.dash = Some(serde_json::from_value(value.clone()).map_err(|_| bad("chart.addElement", "invalid line dash"))?);
                    }
                    if let Some(v) = bool_param(p, "displayEquation") {
                        t.display_equation = v;
                    }
                    if let Some(v) = bool_param(p, "displayRSquared") {
                        t.display_r_squared = v;
                    }
                    if let Some(se) = c.series.get_mut(i)
                        && let Some(slot) = se.trendlines.get_mut(index as usize)
                    {
                        *slot = t;
                    }
                    return Ok(());
                }
                let mut t = trendline_param(p)?;
                t.intercept = f64_param(p, "intercept");
                t.name = str_param(p, "name").map(str::to_owned);
                let eligible: Vec<usize> =
                    targets.into_iter().filter(|i| gridcraft_chart::supports_trendline(gridcraft_chart::series_plot_kind(c, *i))).collect();
                if eligible.is_empty() {
                    return Err(bad("chart.addElement", "trendlines need a 2-D column, bar, line, area, scatter or bubble series"));
                }
                if eligible.iter().any(|i| c.series.get(*i).is_some_and(|se| se.trendlines.len() >= 256)) {
                    return Err(bad("chart.addElement", "trendline budget"));
                }
                for i in eligible {
                    if let Some(se) = c.series.get_mut(i) {
                        se.trendlines.push(t.clone());
                    }
                }
            }
            other => return Err(bad("chart.addElement", format!("unknown element `{other}`"))),
        }
        Ok(())
    })
}

/// The trendline described by `chart.addElement`'s trendline parameters (Excel's Add Chart
/// Element ▸ Trendline choices, plus the Format Trendline options).
fn trendline_param(p: &Json) -> Result<Trendline> {
    let num = |k: &str| Trendline::clamp_forecast(f64_param(p, k).unwrap_or(0.0));
    let mut t = Trendline {
        identity: Default::default(),
        forward: num("forward"),
        backward: num("backward"),
        display_equation: bool_param(p, "displayEquation").unwrap_or(false),
        display_r_squared: bool_param(p, "displayRSquared").unwrap_or(false),
        ..Trendline::default()
    };
    t.kind = match str_param(p, "type").unwrap_or("linear") {
        "linear" => TrendlineKind::Linear,
        // Excel's Linear Forecast: a linear trendline projected two periods ahead.
        "linearForecast" => {
            if t.forward == 0.0 {
                t.forward = 2.0;
            }
            TrendlineKind::Linear
        }
        "exponential" => TrendlineKind::Exponential,
        "logarithmic" => TrendlineKind::Logarithmic,
        "power" => TrendlineKind::Power,
        "polynomial" => TrendlineKind::Polynomial(u32_param(p, "order").unwrap_or(2).clamp(2, 6) as u8),
        "movingAverage" => TrendlineKind::MovingAverage(u32_param(p, "period").unwrap_or(2).clamp(2, 255)),
        other => return Err(bad("chart.addElement", format!("unknown trendline type `{other}`"))),
    };
    Ok(t)
}

pub(super) fn quick_layout(s: &mut Session, p: &Json) -> Result<Json> {
    let n = u32_param(p, "layout").unwrap_or(1);
    with_chart(s, p, |c| {
        let (legend, labels, grid, axes) = match n {
            2 => (LegendPos::Top, true, false, false),
            3 => (LegendPos::Bottom, false, true, false),
            4 => (LegendPos::Bottom, true, false, false),
            5 => (LegendPos::None, false, true, true),
            6 => (LegendPos::Right, false, true, true),
            _ => (LegendPos::Right, false, true, false),
        };
        c.legend = legend;
        c.data_labels = labels;
        c.gridlines = grid;
        if axes {
            c.x_title.get_or_insert_with(|| "Axis Title".into());
            c.y_title.get_or_insert_with(|| "Axis Title".into());
        } else {
            c.x_title = None;
            c.y_title = None;
        }
        Ok(())
    })
}

pub(super) fn change_colors(s: &mut Session, p: &Json) -> Result<Json> {
    let explicit: Option<Vec<Color>> =
        p.get("colors").and_then(Json::as_array).map(|a| a.iter().filter_map(|c| c.as_str().and_then(Color::from_hex)).collect());
    let palette = str_param(p, "palette").unwrap_or("colorful").to_string();
    with_chart(s, p, |c| {
        for (i, se) in c.series.iter_mut().enumerate() {
            se.color = match &explicit {
                Some(cols) if !cols.is_empty() => cols.get(i % cols.len()).copied(),
                _ => match palette.as_str() {
                    "monochrome" => Some(Color::Theme(4, (-400 + (i as i16 * 200)).clamp(-800, 800))),
                    p if p.starts_with("accent") => {
                        let k: u8 = p.trim_start_matches("accent").parse().unwrap_or(1);
                        Some(Color::Theme(3 + k.clamp(1, 6), (-500 + i as i16 * 250).clamp(-800, 800)))
                    }
                    _ => None,
                },
            };
        }
        Ok(())
    })
}

pub(super) fn select_data(s: &mut Session, p: &Json) -> Result<Json> {
    let d = s.doc()?;
    // A new data range rebuilds the series, as in Excel's Select Data › Chart data range.
    let rebuilt = match str_param(p, "range") {
        Some(range) => {
            let range = range.trim_start_matches('=');
            let (sheet_name, body) = split_sheet(range);
            let src_si = match sheet_name {
                Some(n) => d.wb.sheet_index(&n).ok_or_else(|| bad("chart.selectData", "no such sheet"))?,
                None => d.wb.active_sheet(),
            };
            let r = RangeRef::parse(body).ok_or_else(|| bad("chart.selectData", "not a range"))?;
            if r.count() > 1_000_000 {
                return Err(bad("chart.selectData", "range too large for a chart"));
            }
            let src = d.wb.sheet(src_si).ok_or(EngineError::NoDocument)?;
            let by_rows = bool_param(p, "byRows").unwrap_or(r.width() > r.height() + 1);
            let series = crate::cmd::insert::series_from_range(src, r, by_rows);
            Some((series, format!("{}!{}", gridcraft_formula::quote_sheet(&src.name), r.a1()), by_rows))
        }
        None => None,
    };
    // Horizontal (Category) Axis Labels › Edit: one labels range for every series ("" clears).
    let categories = match str_param(p, "categories") {
        None => None,
        Some(t) if t.trim().trim_start_matches('=').is_empty() => Some(None),
        Some(t) => {
            let t = t.trim().trim_start_matches('=');
            let (sheet_name, body) = split_sheet(t);
            let si = match sheet_name {
                Some(n) => d.wb.sheet_index(&n).ok_or_else(|| bad("chart.selectData", "no such sheet"))?,
                None => d.wb.active_sheet(),
            };
            let r = RangeRef::parse(body).ok_or_else(|| bad("chart.selectData", "`categories` is not a range"))?;
            let name = d.wb.sheet(si).map(|sh| sh.name.clone()).ok_or(EngineError::NoDocument)?;
            let abs = |c: CellRef| format!("${}${}", gridcraft_core::col_to_letters(c.col), c.row + 1);
            Some(Some(format!("{}!{}:{}", gridcraft_formula::quote_sheet(&name), abs(r.start), abs(r.end))))
        }
    };
    let mut remove: Vec<usize> = p
        .get("removeSeries")
        .and_then(Json::as_array)
        .map(|a| a.iter().filter_map(Json::as_u64).filter_map(|i| usize::try_from(i).ok()).collect())
        .unwrap_or_default();
    remove.sort_unstable();
    remove.dedup();
    let series_index = u32_param(p, "series").map(|i| i as usize);
    let series_patch = p.get("seriesPatch").cloned();
    let append = bool_param(p, "append").unwrap_or(false);
    if let Some(i) = series_index {
        let (si, ci) = find_chart(s, p)?;
        if s.doc()?.wb.sheet(si).and_then(|sh| sh.charts.get(ci)).is_none_or(|c| i >= c.series.len()) {
            return Err(bad("chart.selectData", "no such series"));
        }
    }
    if let Some(patch) = &series_patch {
        if series_index.is_none() || !patch.is_object() {
            return Err(bad("chart.selectData", "seriesPatch requires a series index and object"));
        }
        if patch.as_object().is_some_and(|map| {
            map.keys().any(|key| {
                !matches!(
                    key.as_str(),
                    "name"
                        | "values"
                        | "categories"
                        | "bubbleSizes"
                        | "smooth"
                        | "secondary"
                        | "type"
                        | "subtype"
                        | "inclusiveQuartiles"
                        | "plotOrder"
                )
            })
        }) {
            return Err(bad("chart.selectData", "unknown series field"));
        }
        for key in ["smooth", "secondary", "inclusiveQuartiles"] {
            if patch.get(key).is_some_and(|v| !v.is_boolean()) {
                return Err(bad("chart.selectData", "series flag must be Boolean"));
            }
        }
        for key in ["type", "subtype"] {
            if patch.get(key).is_some_and(|v| !v.is_string()) {
                return Err(bad("chart.selectData", "series kind must be text"));
            }
        }
        for key in ["name", "values", "categories", "bubbleSizes"] {
            if patch.get(key).is_some_and(|v| (!v.is_null() && !v.is_string()) || v.as_str().is_some_and(|text| text.len() > 65536)) {
                return Err(bad("chart.selectData", "series binding must be text or null"));
            }
        }
        if let Some(ty) = patch.get("type").and_then(Json::as_str)
            && !matches!(
                ty,
                "column"
                    | "bar"
                    | "line"
                    | "pie"
                    | "doughnut"
                    | "area"
                    | "scatter"
                    | "bubble"
                    | "radar"
                    | "histogram"
                    | "waterfall"
                    | "funnel"
                    | "treemap"
                    | "sunburst"
                    | "boxWhisker"
                    | "stock"
                    | "combo"
            )
        {
            return Err(bad("chart.selectData", "unknown chart kind"));
        }
        if let Some(order) = patch.get("plotOrder") {
            let (si, ci) = find_chart(s, p)?;
            let len = s.doc()?.wb.sheet(si).and_then(|sh| sh.charts.get(ci)).map_or(0, |c| c.series.len());
            if order.as_u64().is_none_or(|n| n >= len as u64) {
                return Err(bad("chart.selectData", "plot order bounds"));
            }
        }
        if patch.get("values").is_some_and(|v| !v.is_string()) {
            return Err(bad("chart.selectData", "values must be text"));
        }
    }
    if rebuilt.is_none() && categories.is_none() && remove.is_empty() && series_patch.is_none() {
        return Err(bad("chart.selectData", "missing `range`, `categories` or `removeSeries`"));
    }
    with_chart(s, p, move |c| {
        if let Some((mut series, source, by_rows)) = rebuilt {
            crate::cmd::insert::scatter_x_from_first(c.kind, &mut series);
            if series.is_empty() {
                return Err(bad("chart.selectData", "source has no series"));
            }
            if append {
                if c.series.len().saturating_add(series.len()) > 255 {
                    return Err(bad("chart.selectData", "series budget"));
                }
                c.series.extend(series);
                c.source = None;
            } else {
                if series.len() > 255 {
                    return Err(bad("chart.selectData", "series budget"));
                }
                c.series = series;
                c.source = Some(source);
                c.by_rows = by_rows;
            }
        }
        if let (Some(i), Some(patch)) = (series_index, series_patch) {
            if let Some(ty) = patch.get("type").and_then(Json::as_str) {
                let kind = crate::cmd::insert::chart_kind(ty, patch.get("subtype").and_then(Json::as_str).unwrap_or(""));
                if c.series.len() == 1 {
                    c.kind = kind;
                } else if c.kind != crate::model::ChartKind::Combo {
                    let old = c.kind;
                    for se in &mut c.series {
                        se.kind = Some(old);
                    }
                    c.kind = crate::model::ChartKind::Combo;
                }
            }
            let se = c.series.get_mut(i).ok_or_else(|| bad("chart.selectData", "series disappeared"))?;
            if let Some(v) = patch.get("name") {
                se.name = v.as_str().map(str::to_owned);
            }
            if let Some(v) = patch.get("values").and_then(Json::as_str) {
                se.values = v.to_owned();
            }
            if let Some(v) = patch.get("categories") {
                se.categories = v.as_str().map(str::to_owned);
            }
            if let Some(v) = patch.get("bubbleSizes") {
                se.bubble_sizes = v.as_str().map(str::to_owned);
            }
            if let Some(v) = patch.get("smooth").and_then(Json::as_bool) {
                se.smooth = v;
            }
            if let Some(v) = patch.get("inclusiveQuartiles").and_then(Json::as_bool) {
                se.quartiles = if v { crate::model::QuartileMethod::Inclusive } else { crate::model::QuartileMethod::Exclusive };
            }
            if let Some(v) = patch.get("secondary").and_then(Json::as_bool) {
                se.secondary = v;
            }
            if let Some(v) = patch.get("type").and_then(Json::as_str) {
                se.kind = Some(crate::cmd::insert::chart_kind(v, patch.get("subtype").and_then(Json::as_str).unwrap_or("")));
            }
            if let Some(order) = patch.get("plotOrder").and_then(Json::as_u64).and_then(|n| usize::try_from(n).ok()) {
                let se = c.series.remove(i);
                c.series.insert(order, se);
            }
        }
        for i in remove.iter().rev() {
            if *i < c.series.len() && c.series.len() > 1 {
                c.series.remove(*i);
            }
        }
        if let Some(cats) = categories {
            for se in &mut c.series {
                se.categories = cats.clone();
            }
        }
        Ok(())
    })
}

pub(super) fn move_chart(s: &mut Session, p: &Json) -> Result<Json> {
    let mut owner = p.clone();
    if let Some(map) = owner.as_object_mut() {
        map.remove("sheet");
        if let Some(source) = p.get("sourceSheet") {
            map.insert("sheet".into(), source.clone());
        }
    }
    let (si, ci) = find_chart(s, &owner)?;
    let target = match p.get("sheet").and_then(Json::as_str) {
        Some(n) => Some(s.doc()?.wb.sheet_index(n).ok_or_else(|| bad("chart.move", "no such sheet"))?),
        None => None,
    };
    let at = cell_param_on(s, p, "at", target.unwrap_or(si))?;
    edit(s, |cx| {
        let mut chart = cx.sheet_mut(si)?.charts.remove(ci);
        if let Some(a) = at {
            chart.anchor.cell = a;
        }
        let dest = target.unwrap_or(si);
        cx.sheet_mut(dest)?.charts.push(chart);
        Ok(Json::Null)
    })
}

pub(super) fn format_series(s: &mut Session, p: &Json) -> Result<Json> {
    let idx = u32_param(p, "series").unwrap_or(0) as usize;
    let color = str_param(p, "color").and_then(Color::from_hex);
    with_chart(s, p, |c| {
        if let Some(se) = c.series.get_mut(idx) {
            se.color = color;
        }
        Ok(())
    })
}
