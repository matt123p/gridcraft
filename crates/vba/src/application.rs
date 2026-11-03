//! Application members backed by real engine state. No compatibility-only settings.
use super::objects::{arguments, numeric_index, required};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use serde_json::json;

impl Bridge {
    pub(crate) fn application_state(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        if *object != Object::Application {
            return None;
        }
        let member = call.member.to_ascii_lowercase();
        match member.as_str() {
            "name" | "version" | "operatingsystem" if matches!(call.kind, InvokeKind::Get | InvokeKind::Call) => Some((|| {
                arguments(call, &[])?;
                values::vba_text(match member.as_str() {
                    "name" => "GridCraft",
                    "version" => env!("CARGO_PKG_VERSION"),
                    _ => std::env::consts::OS,
                })
            })()),
            "calculatefull" | "calculatefullrebuild" if call.kind == InvokeKind::Call => Some((|| {
                arguments(call, &[])?;
                self.prepare_application_documents(session)?;
                session.calculate_application().map_err(|e| failure(1004, e.to_string()))?;
                Ok(Value::Empty)
            })()),
            "iteration" | "maxiterations" | "maxchange" => Some((|| {
                if matches!(call.kind, InvokeKind::Get | InvokeKind::Call) {
                    arguments(call, &[])?;
                    let settings = &session.doc().map_err(|e| failure(91, e.to_string()))?.wb.calc;
                    return Ok(match member.as_str() {
                        "iteration" => Value::Boolean(settings.iterative),
                        "maxiterations" => Value::Long(i32::try_from(settings.max_iterations).map_err(|_| failure(6, "MaxIterations overflow"))?),
                        _ => Value::Double(settings.max_change),
                    });
                }
                if call.kind != InvokeKind::Let {
                    return Err(failure(13, "calculation setting requires Let"));
                }
                let args = arguments(call, &["Value"])?;
                let value = required(&args, 0)?;
                let params = match member.as_str() {
                    "iteration" => json!({"iterative": super::address::boolean(value)?}),
                    "maxiterations" => {
                        let count = numeric_index(value)?;
                        if !(1..=32767).contains(&count) {
                            return Err(failure(1004, "MaxIterations must be between 1 and 32767"));
                        }
                        json!({"maxIterations":count})
                    }
                    _ => {
                        let Value::Double(number) = vba_runtime::bind(vba_values::ScalarKind::Double, value).map_err(values::runtime_error)? else {
                            return Err(failure(13, "numeric MaxChange required"));
                        };
                        if !number.is_finite() || number <= 0.0 {
                            return Err(failure(1004, "MaxChange must be finite and positive"));
                        }
                        json!({"maxChange":number})
                    }
                };
                let docs = self.prepare_application_documents(session)?;
                for doc in docs {
                    session
                        .execute_in(gridcraft_engine::Target::Doc(doc), "formulas.calculationOptions", params.clone())
                        .map_err(|e| failure(1004, e.to_string()))?;
                }
                Ok(Value::Empty)
            })()),
            "run" if matches!(call.kind, InvokeKind::Call | InvokeKind::Get) => {
                Some(Err(failure(1004, "Application.Run requires a live project host continuation")))
            }
            _ => None,
        }
    }
    /// Reserve all participating undo groups and validate consent before an app-wide mutation.
    pub(crate) fn prepare_application_documents(&mut self, session: &mut Session) -> HostResult<Vec<u64>> {
        if session.documents().len() > 64 {
            return Err(failure(7, "calculation document budget"));
        }
        let docs = session.documents().iter().map(|doc| doc.uid).collect::<Vec<_>>();
        for doc in &docs {
            let target = session.document(*doc).ok_or_else(|| failure(91, "workbook closed"))?;
            if !target.trust.allows_code() && !(*doc == self.doc && self.consent) {
                return Err(failure(70, "target workbook requires consent"));
            }
        }
        for doc in &docs {
            self.ensure_group(session, *doc).map_err(|e| failure(7, e.to_string()))?;
        }
        Ok(docs)
    }
}
