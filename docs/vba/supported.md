# Supported VBA functions and features

M7A.3 (2026-11-04): native worksheet Button, CheckBox, OptionButton, DropDown,
ListBox, ScrollBar, Spinner, Label and GroupBox controls share the workbook Shape
store. Import, widgets, trusted assignments, live cell/list bindings, selected
Shapes/ControlFormat APIs and native save/reopen are implemented for the
[bounded worksheet profile](worksheet-controls.md).

Updated 2026-11-05: M7's authored `.xlsm` workflows preserve original VBA
project bytes, code names and supported related payload/relationship graphs
through document edits and save/reopen. Loading and compiling are inert; execution
requires document permission. See [M7 acceptance](../../../tiny-vba/docs/m7-acceptance.md)
for exact source/typed-result, cross-workbook context, event and WASM evidence.
Opaque signature payload preservation does not verify signature validity.

This is the feature reference for `GridCraft.Excel.v1`. It describes GridCraft's
VBA integration, including the language runtime supplied by its pinned tiny-vba
dependency. [Getting started](../vba.md) covers build options and the editor.

A supported object member implements the GridCraft behavior described here. It
may accept fewer options than the corresponding Excel member. Use the
[unsupported reference](unsupported.md) for those restrictions. The searchable
[source contracts](reference/vba-object-model-contracts.json) retain signatures,
arguments and implementation/test references.

## Language and editor

| Feature | Available behavior |
|---|---|
| Procedures | Subs and Functions, parameters, return values, optional/default arguments and ParamArray |
| Values | VBA scalar types, Variant, Empty/Null/Missing, CVErr, typed variables and Nothing |
| Arrays | Bounded arrays, declared bounds, indexing and VBA array intrinsics; rectangular worksheet values |
| Objects | Registered Excel/Office nominal types, Object/Variant late binding, properties, With and supported collection defaults/enumeration |
| Modules | Standard, class and document-module source; workbook/sheet context binding and read-only member completion |
| Control flow | Conditions, loops, procedure calls and VBA error handlers within the runtime's resource limits |
| Debugging | Compile diagnostics, breakpoints, pause/resume, step into/over/out, locals and output |
| Worksheet Form Controls | Nine native families, selected ControlFormat APIs, linked data, trusted assignments and native OOXML/VML persistence; [limits](worksheet-controls.md) |
| UserForms | Selected native `.frm`/`.frx` and embedded UserForm import/attachment for flat Label/TextBox/CommandButton/CheckBox/OptionButton/ToggleButton/SpinButton/ScrollBar controls; shared designer/runtime/debugger. Native saving remains unsupported; legacy `.tvform` removal is separate M7A.6 work |

Public no-argument Subs appear in the macro picker. Public Functions become
worksheet functions after source is applied. Private procedures and `Option
Private Module` entries are not public macro-picker entries.

## Built-in VBA functions

These are VBA runtime functions, called directly, for example `Left(text, 3)`.
They are distinct from Excel worksheet functions, such as
`Application.WorksheetFunction.Left(text, 3)`.

| Group | Names | Conditions |
|---|---|---|
| Text slicing and construction | `Left`, `Left$`, `Right`, `Right$`, `Mid`, `Mid$`, `Len`, `Trim`, `Trim$`, `LTrim`, `LTrim$`, `RTrim`, `RTrim$`, `Space`, `Space$`, `String`, `String$`, `StrReverse` | Strings use UTF-16 indexing and bounded allocations |
| Search, comparison and lists | `InStr`, `InStrRev`, `StrComp`, `Replace`, `Split`, `Join`, `Filter` | Binary comparison; Text comparison needs a provider that this integration does not configure |
| Case and character conversion | `LCase`, `LCase$`, `UCase`, `UCase$`, `Asc`, `AscW`, `AscB`, `Chr`, `Chr$`, `ChrW`, `ChrB`, `ChrB$`, `StrConv`, `LenB`, `LeftB`, `LeftB$`, `RightB`, `RightB$`, `MidB`, `MidB$` | Invariant locale and UTF-16 context; code-page/locale-dependent conversions can refuse unsupported inputs |
| Conversion | `CBool`, `CByte`, `CInt`, `CLng`, `CLngLng`, `CLngPtr`, `CSng`, `CDbl`, `CCur`, `CDec`, `CDate`, `CStr`, `CVar`, `Val`, `Str`, `Str$`, `Hex`, `Hex$`, `Oct`, `Oct$`, `CVErr` | Numeric/value conversions are bounded by the target type; locale-dependent conversions and some Decimal inputs are restricted below |
| Inspection | `IsArray`, `IsDate`, `IsEmpty`, `IsError`, `IsMissing`, `IsNull`, `IsNumeric`, `IsObject`, `TypeName`, `VarType` | Inspect VBA values; string-to-number/date recognition depends on configured services |
| Arrays and selection | `Array`, `LBound`, `UBound`, `Choose`, `IIf`, `Switch`, `Partition` | VBA evaluation and argument rules apply; `IIf` does not provide short-circuit evaluation |
| Mathematics | `Abs`, `Atn`, `Cos`, `Exp`, `Fix`, `Int`, `Log`, `Round`, `Sgn`, `Sin`, `Sqr`, `Tan`, `Rnd` | VBA coercion, domain and overflow errors; random state belongs to the execution session |
| Financial | `DDB`, `FV`, `IPmt`, `IRR`, `MIRR`, `NPer`, `NPV`, `Pmt`, `PPmt`, `PV`, `Rate`, `SLN`, `SYD` | Native runtime financial functions, with explicit invalid-domain/convergence errors |
| Date arithmetic | `DateAdd`, `DateDiff`, `DatePart`, `DateSerial`, `TimeSerial`, `Day`, `Month`, `Year`, `Hour`, `Minute`, `Second`, `Weekday` | Gregorian numeric-date operations; locale text parsing and alternate calendars are not supplied |
| Color and invocation | `RGB`, `QBColor`, `CallByName` | Dynamic invocation remains subject to the supported object graph and host-call policy |

`Randomize` with an explicit seed and `Debug.Print` are also available.
The runtime keeps fixed-string `$` spellings distinct where listed.

Numeric Date/Currency values retain their VBA types inside execution. `CLngPtr`
uses the target's pointer width. `CDec` supports exact Decimal/integral/Boolean/
Empty/Currency conversions; floating-point/Date conversion and locale-dependent
string conversion have additional restrictions. Clock reads, localized formatting
and other service-dependent intrinsics are listed in
[unsupported runtime services](unsupported.md#runtime-functions-that-need-a-service).

## Worksheet functions

Use `Application.WorksheetFunction.Name(...)` for the calculator-backed function
bridge. A Range argument retains its reference context; bounded scalar and array
arguments are also accepted. An error returned by the calculator becomes VBA
error 1004. Calls accept up to 30 arguments; the explicit member signatures define
required and optional slots. Underscores in VBA names map to dots in calculator
names, for example `Norm_Dist` maps to `NORM.DIST`.

### Functions with declared member signatures

<!-- BEGIN VBA WORKSHEET FUNCTIONS -->

- `AccrInt`, `AccrIntM`, `Acos`, `Acosh`, `Aggregate`, `AmorDegrc`, `AmorLinc`, `And`
- `Asc`, `Asin`, `Asinh`, `Atan2`, `Atanh`, `AveDev`, `Average`, `AverageIf`
- `AverageIfs`, `BahtText`, `BesselI`, `BesselJ`, `BesselK`, `BesselY`, `BetaDist`, `BetaInv`
- `Beta_Dist`, `Beta_Inv`, `Bin2Dec`, `Bin2Hex`, `Bin2Oct`, `BinomDist`, `Binom_Dist`, `Binom_Inv`
- `Ceiling`, `Ceiling_Precise`, `ChiDist`, `ChiInv`, `ChiSq_Dist`, `ChiSq_Dist_RT`, `ChiSq_Inv`, `ChiSq_Inv_RT`
- `ChiSq_Test`, `ChiTest`, `Choose`, `Clean`, `Combin`, `Complex`, `Confidence`, `Confidence_Norm`
- `Confidence_T`, `Convert`, `Correl`, `Cosh`, `Count`, `CountA`, `CountBlank`, `CountIf`
- `CountIfs`, `CoupDayBs`, `CoupDays`, `CoupDaysNc`, `CoupNcd`, `CoupNum`, `CoupPcd`, `Covar`
- `Covariance_P`, `Covariance_S`, `CritBinom`, `CumIPmt`, `CumPrinc`, `DAverage`, `DCount`, `DCountA`
- `DGet`, `DMax`, `DMin`, `DProduct`, `DStDev`, `DStDevP`, `DSum`, `DVar`
- `DVarP`, `Days360`, `Db`, `Dbcs`, `Ddb`, `Dec2Bin`, `Dec2Hex`, `Dec2Oct`
- `Degrees`, `Delta`, `DevSq`, `Disc`, `Dollar`, `DollarDe`, `DollarFr`, `Duration`
- `EDate`, `Effect`, `EoMonth`, `Erf`, `ErfC`, `ErfC_Precise`, `Erf_Precise`, `Even`
- `ExponDist`, `Expon_Dist`, `FDist`, `FInv`, `FTest`, `FVSchedule`, `F_Dist`, `F_Dist_RT`
- `F_Inv`, `F_Inv_RT`, `F_Test`, `Fact`, `FactDouble`, `Find`, `FindB`, `Fisher`
- `FisherInv`, `Fixed`, `Floor`, `Floor_Precise`, `Forecast`, `Frequency`, `Fv`, `GammaDist`
- `GammaInv`, `GammaLn`, `GammaLn_Precise`, `Gamma_Dist`, `Gamma_Inv`, `Gcd`, `GeStep`, `GeoMean`
- `Growth`, `HLookup`, `HarMean`, `Hex2Bin`, `Hex2Dec`, `Hex2Oct`, `HypGeomDist`, `HypGeom_Dist`
- `ISO_Ceiling`, `IfError`, `ImAbs`, `ImArgument`, `ImConjugate`, `ImCos`, `ImDiv`, `ImExp`
- `ImLn`, `ImLog10`, `ImLog2`, `ImPower`, `ImProduct`, `ImReal`, `ImSin`, `ImSqrt`
- `ImSub`, `ImSum`, `Imaginary`, `Index`, `IntRate`, `Intercept`, `Ipmt`, `Irr`
- `IsErr`, `IsError`, `IsEven`, `IsLogical`, `IsNA`, `IsNonText`, `IsNumber`, `IsOdd`
- `IsText`, `Ispmt`, `Kurt`, `Large`, `Lcm`, `LinEst`, `Ln`, `Log`
- `Log10`, `LogEst`, `LogInv`, `LogNormDist`, `LogNorm_Dist`, `LogNorm_Inv`, `Lookup`, `MDeterm`
- `MDuration`, `MInverse`, `MIrr`, `MMult`, `MRound`, `Match`, `Max`, `Median`
- `Min`, `Mode`, `Mode_Mult`, `Mode_Sngl`, `MultiNomial`, `NPer`, `NegBinomDist`, `NegBinom_Dist`
- `NetworkDays`, `NetworkDays_Intl`, `Nominal`, `NormDist`, `NormInv`, `NormSDist`, `NormSInv`, `Norm_Dist`
- `Norm_Inv`, `Norm_S_Dist`, `Norm_S_Inv`, `Npv`, `Oct2Bin`, `Oct2Dec`, `Oct2Hex`, `Odd`
- `OddFPrice`, `OddFYield`, `OddLPrice`, `OddLYield`, `Or`, `Pearson`, `PercentRank`, `PercentRank_Exc`
- `PercentRank_Inc`, `Percentile`, `Percentile_Exc`, `Percentile_Inc`, `Permut`, `Phonetic`, `Pi`, `Pmt`
- `Poisson`, `Poisson_Dist`, `Power`, `Ppmt`, `Price`, `PriceDisc`, `PriceMat`, `Prob`
- `Product`, `Proper`, `Pv`, `Quartile`, `Quartile_Exc`, `Quartile_Inc`, `Quotient`, `RSq`
- `Radians`, `RandBetween`, `Rank`, `Rank_Avg`, `Rank_Eq`, `Rate`, `Received`, `Replace`
- `ReplaceB`, `Rept`, `Roman`, `Round`, `RoundDown`, `RoundUp`, `Search`, `SearchB`
- `SeriesSum`, `Sinh`, `Skew`, `Sln`, `Slope`, `Small`, `SqrtPi`, `StDev`
- `StDevP`, `StDev_P`, `StDev_S`, `StEyx`, `Standardize`, `Substitute`, `Subtotal`, `Sum`
- `SumIf`, `SumIfs`, `SumProduct`, `SumSq`, `SumX2MY2`, `SumX2PY2`, `SumXMY2`, `Syd`
- `TBillEq`, `TBillPrice`, `TBillYield`, `TDist`, `TInv`, `TTest`, `T_Dist`, `T_Dist_2T`
- `T_Dist_RT`, `T_Inv`, `T_Inv_2T`, `T_Test`, `Tanh`, `Text`, `Transpose`, `Trend`
- `Trim`, `TrimMean`, `USDollar`, `VLookup`, `Var`, `VarP`, `Var_P`, `Var_S`
- `Vdb`, `WeekNum`, `Weekday`, `Weibull`, `Weibull_Dist`, `WorkDay`, `WorkDay_Intl`, `Xirr`
- `Xnpv`, `YearFrac`, `YieldDisc`, `YieldMat`, `ZTest`, `Z_Test`

<!-- END VBA WORKSHEET FUNCTIONS -->

### Additional functions through generic dispatch

The following functions use the existing calculator-dispatch route. They do not
yet have individual typed member signatures in the adapter catalogue. Treat their
arguments as the calculator contract and do not assume Excel-specific named
parameter metadata or complete Excel behavior.

<!-- BEGIN VBA GENERIC FUNCTIONS -->

- `acot`, `acoth`, `arabic`, `base`, `binom_dist_range`, `bitand`, `bitlshift`, `bitor`
- `bitrshift`, `bitxor`, `ceiling_math`, `combina`, `cot`, `coth`, `csc`, `csch`
- `days`, `decimal`, `encodeurl`, `floor_math`, `forecast_ets`, `forecast_ets_confint`, `forecast_ets_seasonality`, `forecast_ets_stat`
- `forecast_linear`, `gamma`, `gauss`, `ifna`, `imcosh`, `imcot`, `imcsc`, `imcsch`
- `imsec`, `imsech`, `imsinh`, `imtan`, `isformula`, `isoweeknum`, `munit`, `numbervalue`
- `pduration`, `permutationa`, `phi`, `rri`, `sec`, `sech`, `skew_p`, `unichar`
- `unicode`, `xor`

<!-- END VBA GENERIC FUNCTIONS -->

`WEBSERVICE`, `FILTERXML`, CUBE provider functions and `RTD` are not working
external-data services; see [unsupported functions](unsupported.md#worksheet-functions).

## Workbook and worksheet features

| Area | Supported features | Main boundary |
|---|---|---|
| Application and workbooks | Live Workbooks collection, code-owner versus active workbook, activation, application identity, native calculation and iteration settings, current-project `Application.Run` | Run stays in the current project; calculation follows native scope/settings |
| Files | Add/Open, Save/SaveAs/SaveCopyAs/Close, native OOXML output, password-assisted opening and host-provided file/save decisions | Save options are bounded; opening never grants code consent |
| Worksheets | Add, same-workbook copy/move/delete, name/index/code name, visibility, activation and selection | Unsupported project/package transfers refuse explicitly |
| Ranges | A1/named/qualified/multi-area references, Cells/Rows/Columns, Areas, Item/default access, Offset/Resize, End, Find navigation and native Insert/Delete | Ownership, scan and materialization limits apply |
| Values and formulas | Value, Value2, default Value, Formula/Formula2, array and A1/R1C1/local formula variants, Address, ClearContents | Value and Value2 differ for Date/Currency; worksheet storage is numeric/text/error state |
| Clipboard and editing | Native Copy/Cut/Paste, selected Range.PasteSpecial options, fills, Replace and clear operations | Bounded same-workbook clipboard; no system clipboard or arbitrary paste modes |
| Formatting | Number formats, fonts, fills, borders, alignment, row/column sizing, native named styles and represented conditional-format rules | Only native stored fields and supported option combinations |
| Data tools | Sorting, AutoFilter/filter criteria, validation, outline/grouping, native what-if/scenarios and represented table operations | Provider, model and unsupported sort/converter options are excluded |
| Names and evaluation | Workbook/worksheet Names, defined-name lifetime, Evaluate and calculator-backed WorksheetFunction calls | Native reference/array/error semantics; no automatic compatibility fallback |
| Structured tables | Local ListObjects, columns/rows, native names, totals and supported range/style operations | No SharePoint/XML/import-provider graph |
| Charts and drawings | Existing embedded charts, native series/trendlines/axes/titles/labels, chart areas, shapes, pictures and text/fill/line formatting | No chart sheets, arbitrary effects or invented child state |
| Pivots and slicers | Local pivot reports/fields/items, supported filters/calculated fields, report layout and existing table slicers | No external cache/provider/data-model lifecycle |
| Windows and navigation | Existing native Window/Windows identities, per-window active sheet/selection, supported Goto/scroll/history operations, saved views and watches | Several Excel independent-pane/view/geometry members remain unavailable |
| Protection | Native sheet/workbook passwords and permissions, protected selection, Range.AllowEdit and AllowEditRanges | No UI-only protection, IRM, resolved ACL or encryption-writing service |
| Printing and export | Native PageSetup, headers/footers, print area/titles, page/break snapshots, PDF export and connected print/preview services | Options, scopes and transport must be available on the host |
| Annotations | Native notes/threaded replies, cell hyperlinks and six built-in document properties | Local content and internal navigation; no cloud identity or arbitrary property graph |
| Interaction | Native dialog/picker continuations, filename pickers, FindFile, status-bar visibility, column autocomplete and cell checkboxes | Only existing dialogs/controls with an explicit host implementation |

Date-formatted Value reads return VBA Date; Currency-formatted reads return
Currency. Value2 retains numeric values. Multi-cell reads return one-based 2D
Variant arrays. Writes accept scalars or compatible 2D arrays and use native
protection, recalculation and undo.

## Events and worksheet functions written in VBA

The event profile includes `Workbook_Open`, `Worksheet_Change`, workbook
BeforeSave/BeforeClose/AfterSave and the corresponding Application
WorkbookBeforeSave/WorkbookBeforeClose/WorkbookAfterSave callbacks. BeforeSave
and BeforeClose carry cancellation before the native effect. AfterSave follows
actual I/O. `EnableEvents` controls this bounded queue.

Worksheet_Change covers the connected native edit operations and VBA value
writes, clears, paste, Replace and fills. Formatting-only operations do not create
value-change events. Events are queued rather than delivered reentrantly. See
[events and cancellation](integration.md#events-and-cancellation) for ordering.

Applied public VBA Functions run as read-only worksheet UDFs. Supported parameters
include scalars, Variant arrays and read-only Range/Object arguments. Results can
be scalars or rectangular spilled arrays. The UDF context supports workbook value
reads, a bounded Range API, `Application.Caller` and `Application.Volatile`.
Built-in calculator functions take precedence over UDF names.

## Object-member index

This index is generated from implemented adapter descriptors. It lists member
names, not every supported option or invocation context. Read the feature limits
above and the [restrictions](unsupported.md) before using an Excel macro unchanged.
A listed object is a projection of existing native state, not a complete Excel class.

<!-- BEGIN VBA OBJECT MEMBERS -->

<details>
<summary>Full implemented object member names</summary>

| Object | Members |
|---|---|
| `AboveAverage` | `AboveBelow`, `Application`, `AppliesTo`, `Delete`, `Font`, `Interior`, `ModifyAppliesToRange`, `NumStdDev`, `NumberFormat`, `Parent`, `Priority`, `SetFirstPriority`, `SetLastPriority`, `StopIfTrue`, `Type` |
| `AllowEditRange` | `ChangePassword`, `Delete`, `Range`, `Title`, `Unprotect` |
| `AllowEditRanges` | `Add`, `Count`, `Item` |
| `Application` | `ActiveCell`, `ActivePrinter`, `ActiveSheet`, `ActiveWindow`, `ActiveWorkbook`, `Application`, `Calculate`, `CalculateFull`, `CalculateFullRebuild`, `Calculation`, `Caller`, `Cells`, `Columns`, `CutCopyMode`, `Dialogs`, `DisplayStatusBar`, `EnableEvents`, `Evaluate`, `FindFile`, `GetOpenFilename`, `GetSaveAsFilename`, `Goto`, `Intersect`, `Iteration`, `MaxChange`, `MaxIterations`, `Name`, `Names`, `OperatingSystem`, `Parent`, `PreviousSelections`, `Range`, `Rows`, `Run`, `Selection`, `Sheets`, `ThisWorkbook`, `Union`, `Version`, `Watches`, `Windows`, `Workbooks`, `WorksheetFunction`, `Worksheets` |
| `Areas` | `Application`, `Count`, `Item`, `Parent` |
| `Author` | `Application`, `Name` |
| `AutoFilter` | `Application`, `ApplyFilter`, `FilterMode`, `Filters`, `Parent`, `Range`, `ShowAllData` |
| `Axes` | `Application`, `Count`, `Item`, `Parent` |
| `Axis` | `Application`, `AxisGroup`, `AxisTitle`, `HasMajorGridlines`, `HasTitle`, `MajorGridlines`, `MajorUnit`, `MajorUnitIsAuto`, `MaximumScale`, `MaximumScaleIsAuto`, `MinimumScale`, `MinimumScaleIsAuto`, `Parent`, `Type` |
| `AxisTitle` | `Application`, `Caption`, `Delete`, `Parent`, `Text` |
| `Border` | `Application`, `Color`, `LineStyle`, `Parent`, `Weight` |
| `Borders` | `Application`, `Color`, `Count`, `Item`, `LineStyle`, `Parent`, `Weight` |
| `CalculatedFields` | `Add`, `Application`, `Count`, `Item`, `Parent` |
| `CellControl` | `SetCheckbox`, `Type` |
| `Characters` | `Application`, `Count`, `Delete`, `Font`, `Parent`, `Text` |
| `Chart` | `Application`, `ApplyDataLabels`, `Axes`, `ChartArea`, `ChartStyle`, `ChartTitle`, `ChartType`, `Delete`, `HasAxis`, `HasLegend`, `HasTitle`, `Legend`, `Location`, `Name`, `Parent`, `PivotLayout`, `PlotArea`, `PlotBy`, `Refresh`, `SeriesCollection`, `SetSourceData` |
| `ChartArea` | `Application`, `Height`, `Left`, `Parent`, `Top`, `Width` |
| `ChartFormat` | `Application`, `Fill`, `Line`, `Parent` |
| `ChartObject` | `Application`, `BottomRightCell`, `Chart`, `Delete`, `Height`, `Index`, `Left`, `Name`, `Parent`, `Placement`, `ShapeRange`, `Top`, `TopLeftCell`, `Width`, `ZOrder` |
| `ChartObjects` | `Application`, `Count`, `Delete`, `Item`, `Parent`, `ShapeRange` |
| `ChartTitle` | `Application`, `Caption`, `Delete`, `Parent`, `Text` |
| `ColorFormat` | `Application`, `ObjectThemeColor`, `Parent`, `RGB`, `TintAndShade` |
| `ColorScale` | `Application`, `AppliesTo`, `ColorScaleCriteria`, `Delete`, `ModifyAppliesToRange`, `Parent`, `Priority`, `SetFirstPriority`, `SetLastPriority`, `StopIfTrue`, `Type` |
| `ColorScaleCriteria` | `Count`, `Item` |
| `ColorScaleCriterion` | `FormatColor`, `Index`, `Type`, `Value` |
| `Comment` | `Application`, `Author`, `Delete`, `Next`, `Parent`, `Previous`, `Text`, `Visible` |
| `CommentThreaded` | `AddReply`, `Application`, `Author`, `Delete`, `Next`, `Parent`, `Previous`, `Replies`, `Text` |
| `Comments` | `Application`, `Count`, `Item`, `Parent` |
| `CommentsThreaded` | `Application`, `Count`, `Item`, `Parent` |
| `ConditionValue` | `Application`, `Modify`, `Parent`, `Type`, `Value` |
| `ControlFormat` | `AddItem`, `Application`, `DropDownLines`, `Enabled`, `LargeChange`, `LinkedCell`, `List`, `ListCount`, `ListFillRange`, `ListIndex`, `Max`, `Min`, `MultiSelect`, `Parent`, `PrintObject`, `RemoveAllItems`, `RemoveItem`, `SmallChange`, `Value` |
| `CustomView` | `Application`, `Delete`, `Name`, `Parent`, `PrintSettings`, `RowColSettings`, `Show` |
| `CustomViews` | `Add`, `Application`, `Count`, `Item`, `Parent` |
| `DataLabel` | `Application`, `Caption`, `Parent`, `ShowValue`, `Text` |
| `DataLabels` | `Application`, `Count`, `Item`, `Parent` |
| `Databar` | `Application`, `AppliesTo`, `BarColor`, `Delete`, `MaxPoint`, `MinPoint`, `ModifyAppliesToRange`, `Parent`, `Priority`, `SetFirstPriority`, `SetLastPriority`, `ShowValue`, `StopIfTrue`, `Type` |
| `Dialog` | `Application`, `Parent`, `Show` |
| `Dialogs` | `Application`, `Count`, `Item`, `Parent` |
| `DocumentProperties` | `Application`, `Count`, `Item`, `Parent` |
| `DocumentProperty` | `Application`, `LinkToContent`, `Name`, `Parent`, `Type`, `Value` |
| `FillFormat` | `Application`, `ForeColor`, `Parent`, `Solid`, `Type`, `Visible` |
| `Filter` | `Application`, `Criteria1`, `Criteria2`, `On`, `Operator`, `Parent` |
| `Filters` | `Application`, `Count`, `Item`, `Parent` |
| `Font` | `Application`, `Bold`, `Color`, `ColorIndex`, `Italic`, `Name`, `Parent`, `Size`, `Strikethrough`, `Subscript`, `Superscript`, `ThemeColor`, `TintAndShade`, `Underline` |
| `Font2` | `Application`, `Bold`, `Fill`, `Italic`, `Line`, `Name`, `Parent`, `Size` |
| `FormatColor` | `Application`, `Color`, `Parent`, `ThemeColor`, `TintAndShade` |
| `FormatCondition` | `Application`, `AppliesTo`, `Delete`, `Font`, `Formula1`, `Formula2`, `Interior`, `Modify`, `ModifyAppliesToRange`, `NumberFormat`, `Operator`, `Parent`, `Priority`, `SetFirstPriority`, `SetLastPriority`, `StopIfTrue`, `Type` |
| `FormatConditions` | `Add`, `AddAboveAverage`, `AddColorScale`, `AddDatabar`, `AddIconSetCondition`, `AddTop10`, `AddUniqueValues`, `Application`, `Count`, `Delete`, `Item`, `Parent` |
| `Gridlines` | `Application`, `Parent` |
| `HPageBreak` | `Application`, `Delete`, `Extent`, `Location`, `Parent`, `Type` |
| `HPageBreaks` | `Add`, `Application`, `Count`, `Item`, `Parent` |
| `HeaderFooter` | `Text` |
| `Hyperlink` | `Address`, `Application`, `Delete`, `Follow`, `Parent`, `Range`, `ScreenTip`, `SubAddress`, `TextToDisplay`, `Type` |
| `Hyperlinks` | `Add`, `Application`, `Count`, `Delete`, `Item`, `Parent` |
| `IconCriteria` | `Count`, `Item` |
| `IconCriterion` | `Index`, `Operator`, `Type`, `Value` |
| `IconSet` | `Application`, `Count`, `ID`, `Parent` |
| `IconSetCondition` | `Application`, `AppliesTo`, `Delete`, `IconCriteria`, `IconSet`, `ModifyAppliesToRange`, `Parent`, `Priority`, `ReverseOrder`, `SetFirstPriority`, `SetLastPriority`, `ShowIconOnly`, `StopIfTrue`, `Type` |
| `IconSets` | `Application`, `Count`, `Item`, `Parent` |
| `Interior` | `Application`, `Color`, `ColorIndex`, `Parent`, `Pattern`, `PatternColor`, `PatternColorIndex`, `ThemeColor`, `TintAndShade` |
| `Legend` | `Application`, `Delete`, `Parent`, `Position` |
| `LineFormat` | `Application`, `DashStyle`, `ForeColor`, `Parent`, `Visible`, `Weight` |
| `ListColumn` | `Application`, `DataBodyRange`, `Delete`, `Index`, `Name`, `Parent`, `Range`, `Total`, `TotalsCalculation` |
| `ListColumns` | `Add`, `Application`, `Count`, `Item`, `Parent` |
| `ListObject` | `Application`, `AutoFilter`, `DataBodyRange`, `Delete`, `DisplayName`, `HeaderRowRange`, `ListColumns`, `ListRows`, `Name`, `Parent`, `Range`, `Resize`, `ShowAutoFilter`, `ShowHeaders`, `ShowTableStyleColumnStripes`, `ShowTableStyleFirstColumn`, `ShowTableStyleLastColumn`, `ShowTableStyleRowStripes`, `ShowTotals`, `Sort`, `SourceType`, `TableStyle`, `TotalsRowRange`, `Unlist` |
| `ListObjects` | `Add`, `Application`, `Count`, `Item`, `Parent` |
| `ListRow` | `Application`, `Delete`, `Index`, `Parent`, `Range` |
| `ListRows` | `Add`, `Application`, `Count`, `Item`, `Parent` |
| `Name` | `Application`, `Comment`, `Delete`, `Name`, `NameLocal`, `Parent`, `RefersTo`, `RefersToLocal`, `RefersToR1C1`, `RefersToR1C1Local`, `RefersToRange`, `Value`, `Visible` |
| `Names` | `Add`, `Application`, `Count`, `Item`, `Parent` |
| `Outline` | `Application`, `Parent`, `ShowLevels`, `SummaryColumn`, `SummaryRow` |
| `Page` | `CenterFooter`, `CenterHeader`, `LeftFooter`, `LeftHeader`, `RightFooter`, `RightHeader` |
| `PageSetup` | `AlignMarginsHeaderFooter`, `Application`, `BottomMargin`, `CenterFooter`, `CenterHeader`, `CenterHorizontally`, `CenterVertically`, `DifferentFirstPageHeaderFooter`, `EvenPage`, `FirstPage`, `FitToPagesTall`, `FitToPagesWide`, `FooterMargin`, `HeaderMargin`, `LeftFooter`, `LeftHeader`, `LeftMargin`, `OddAndEvenPagesHeaderFooter`, `Order`, `Orientation`, `Pages`, `PaperSize`, `Parent`, `PrintArea`, `PrintGridlines`, `PrintHeadings`, `PrintTitleColumns`, `PrintTitleRows`, `RightFooter`, `RightHeader`, `RightMargin`, `ScaleWithDocHeaderFooter`, `TopMargin`, `Zoom` |
| `Pages` | `Count`, `Item` |
| `PivotCache` | `Application`, `Index`, `OLAP`, `Parent`, `RecordCount`, `Refresh`, `SourceData`, `SourceType` |
| `PivotCaches` | `Application`, `Count`, `Item`, `Parent` |
| `PivotCell` | `Application`, `Parent`, `PivotCellType`, `PivotTable`, `Range` |
| `PivotField` | `Application`, `AutoSort`, `Calculation`, `Caption`, `ClearAllFilters`, `ClearManualFilter`, `CurrentPage`, `Delete`, `Formula`, `Function`, `HiddenItems`, `IsCalculated`, `Name`, `NumberFormat`, `Orientation`, `Parent`, `PivotItems`, `Position`, `ShowDetail`, `SourceCaption`, `SourceName`, `Value`, `VisibleItems` |
| `PivotFields` | `Application`, `Count`, `Item`, `Parent` |
| `PivotItem` | `Application`, `Caption`, `Name`, `Parent`, `Position`, `RecordCount`, `ShowDetail`, `SourceName`, `Value`, `Visible` |
| `PivotItems` | `Application`, `Count`, `Item`, `Parent` |
| `PivotLayout` | `Application`, `Parent`, `PivotTable` |
| `PivotTable` | `AddDataField`, `AddFields`, `Application`, `CacheIndex`, `CalculatedFields`, `ClearAllFilters`, `ClearTable`, `ColumnFields`, `ColumnGrand`, `DataBodyRange`, `DataFields`, `DataPivotField`, `DisplayFieldCaptions`, `EnableDrilldown`, `HiddenFields`, `Location`, `Name`, `PageFields`, `Parent`, `PivotCache`, `PivotChart`, `PivotFields`, `RefreshTable`, `RowAxisLayout`, `RowFields`, `RowGrand`, `SourceData`, `SubtotalLocation`, `TableRange1`, `TableRange2`, `TableStyle2`, `Update` |
| `PivotTables` | `Add`, `Application`, `Count`, `Item`, `Parent` |
| `PlotArea` | `Application`, `Height`, `InsideHeight`, `InsideLeft`, `InsideTop`, `InsideWidth`, `Left`, `Parent`, `Top`, `Width` |
| `Protection` | `AllowDeletingColumns`, `AllowDeletingRows`, `AllowEditRanges`, `AllowFiltering`, `AllowFormattingCells`, `AllowFormattingColumns`, `AllowFormattingRows`, `AllowInsertingColumns`, `AllowInsertingHyperlinks`, `AllowInsertingRows`, `AllowSorting`, `AllowUsingPivotTables` |
| `Range` | `Activate`, `AddComment`, `AddCommentThreaded`, `Address`, `AdvancedFilter`, `AllowEdit`, `Application`, `Areas`, `AutoComplete`, `AutoFill`, `AutoFilter`, `AutoFit`, `Borders`, `CellControl`, `Cells`, `Clear`, `ClearComments`, `ClearContents`, `ClearFormats`, `ClearHyperlinks`, `ClearNotes`, `Column`, `ColumnWidth`, `Columns`, `Comment`, `CommentThreaded`, `Consolidate`, `Copy`, `Count`, `CountLarge`, `CurrentArray`, `CurrentRegion`, `Cut`, `DataSeries`, `Delete`, `End`, `EntireColumn`, `EntireRow`, `Evaluate`, `ExportAsFixedFormat`, `FillDown`, `FillLeft`, `FillRight`, `FillUp`, `Find`, `FindNext`, `FindPrevious`, `Font`, `FormatConditions`, `Formula`, `Formula2`, `Formula2Local`, `Formula2R1C1`, `Formula2R1C1Local`, `FormulaArray`, `FormulaHidden`, `FormulaLocal`, `FormulaR1C1`, `FormulaR1C1Local`, `GoalSeek`, `Group`, `HasArray`, `HasFormula`, `HasSpill`, `Height`, `Hidden`, `HorizontalAlignment`, `Hyperlinks`, `IndentLevel`, `Insert`, `Interior`, `Item`, `ListObject`, `Locked`, `Merge`, `MergeArea`, `MergeCells`, `Next`, `NoteText`, `NumberFormat`, `NumberFormatLocal`, `Offset`, `Orientation`, `OutlineLevel`, `PageBreak`, `Parent`, `PasteSpecial`, `PivotCell`, `PivotTable`, `Previous`, `PrintOut`, `Range`, `RemoveDuplicates`, `Replace`, `Resize`, `Row`, `RowHeight`, `Rows`, `Select`, `ShowDetail`, `ShrinkToFit`, `Sort`, `SpecialCells`, `SpillParent`, `Style`, `Subtotal`, `Table`, `TextToColumns`, `UnMerge`, `Ungroup`, `Validation`, `Value`, `Value2`, `VerticalAlignment`, `Width`, `Worksheet`, `WrapText` |
| `Scenario` | `Application`, `ChangeScenario`, `ChangingCells`, `Comment`, `Delete`, `Index`, `Name`, `Parent`, `Show`, `Values` |
| `Scenarios` | `Add`, `Application`, `Count`, `CreateSummary`, `Item`, `Parent` |
| `Series` | `Application`, `ApplyDataLabels`, `AxisGroup`, `BubbleSizes`, `ChartType`, `DataLabels`, `Delete`, `Format`, `Formula`, `HasDataLabels`, `Name`, `Parent`, `PlotOrder`, `QuartileCalculationInclusiveMedian`, `Smooth`, `Trendlines`, `Values`, `XValues` |
| `SeriesCollection` | `Add`, `Application`, `Count`, `Item`, `Parent` |
| `Shape` | `AlternativeText`, `Application`, `AutoShapeType`, `BottomRightCell`, `Chart`, `ControlFormat`, `Delete`, `Fill`, `FormControlType`, `HasChart`, `Height`, `ID`, `IncrementLeft`, `IncrementTop`, `Left`, `Line`, `Name`, `OnAction`, `Parent`, `Placement`, `TextFrame`, `TextFrame2`, `Top`, `TopLeftCell`, `Type`, `Visible`, `Width`, `ZOrder`, `ZOrderPosition` |
| `ShapeRange` | `Align`, `Application`, `AutoShapeType`, `Count`, `Delete`, `Fill`, `Height`, `Item`, `Left`, `Line`, `Name`, `Parent`, `TextFrame`, `TextFrame2`, `Top`, `Type`, `Width`, `ZOrder`, `ZOrderPosition` |
| `Shapes` | `AddFormControl`, `AddLine`, `AddPicture`, `AddShape`, `AddTextbox`, `Application`, `Count`, `Item`, `Parent`, `Range` |
| `Sheets` | `Add`, `Application`, `Count`, `Item`, `Parent`, `PrintOut`, `PrintPreview` |
| `Slicer` | `Application`, `Caption`, `Delete`, `Height`, `Left`, `Name`, `NumberOfColumns`, `Shape`, `SlicerCache`, `Top`, `Width` |
| `SlicerCache` | `Application`, `ClearAllFilters`, `ClearManualFilter`, `Delete`, `FilterCleared`, `Index`, `ListObject`, `Name`, `OLAP`, `Parent`, `SlicerCacheType`, `SlicerItems`, `Slicers`, `SourceName`, `SourceType` |
| `SlicerCaches` | `Application`, `Count`, `Item`, `Parent` |
| `SlicerItem` | `Application`, `Caption`, `Name`, `Parent`, `Selected`, `SourceName`, `Value` |
| `SlicerItems` | `Application`, `Count`, `Item`, `Parent` |
| `Slicers` | `Application`, `Count`, `Item`, `Parent` |
| `Sort` | `Application`, `Apply`, `Header`, `MatchCase`, `Orientation`, `Parent`, `SetRange`, `SortFields` |
| `SortField` | `Application`, `CustomOrder`, `DataOption`, `Delete`, `Key`, `Order`, `Parent`, `SortOn` |
| `SortFields` | `Add`, `Add2`, `Application`, `Clear`, `Count`, `Item`, `Parent` |
| `Style` | `Application`, `BuiltIn`, `FormulaHidden`, `Locked`, `Name`, `NameLocal`, `NumberFormat`, `Parent`, `Value` |
| `Styles` | `Add`, `Application`, `Count`, `Item`, `Parent` |
| `TextFrame` | `Application`, `Characters`, `HorizontalAlignment`, `Parent`, `ReadingOrder`, `VerticalAlignment` |
| `TextFrame2` | `Application`, `HasText`, `Parent`, `TextRange`, `VerticalAnchor` |
| `TextRange2` | `Application`, `Delete`, `Font`, `Length`, `Parent`, `Text` |
| `Top10` | `Application`, `AppliesTo`, `Delete`, `Font`, `Interior`, `ModifyAppliesToRange`, `NumberFormat`, `Parent`, `Percent`, `Priority`, `Rank`, `SetFirstPriority`, `SetLastPriority`, `StopIfTrue`, `TopBottom`, `Type` |
| `Trendline` | `Application`, `Backward2`, `Border`, `Delete`, `DisplayEquation`, `DisplayRSquared`, `Format`, `Forward2`, `Index`, `Intercept`, `InterceptIsAuto`, `Name`, `NameIsAuto`, `Order`, `Parent`, `Period`, `Type` |
| `Trendlines` | `Add`, `Application`, `Count`, `Item`, `Parent` |
| `UniqueValues` | `Application`, `AppliesTo`, `Delete`, `DupeUnique`, `Font`, `Interior`, `ModifyAppliesToRange`, `NumberFormat`, `Parent`, `Priority`, `SetFirstPriority`, `SetLastPriority`, `StopIfTrue`, `Type` |
| `VPageBreak` | `Application`, `Delete`, `Extent`, `Location`, `Parent`, `Type` |
| `VPageBreaks` | `Add`, `Application`, `Count`, `Item`, `Parent` |
| `Validation` | `Add`, `AlertStyle`, `Application`, `Delete`, `ErrorMessage`, `ErrorTitle`, `Formula1`, `Formula2`, `IgnoreBlank`, `InCellDropdown`, `InputMessage`, `InputTitle`, `Modify`, `Operator`, `Parent`, `ShowError`, `ShowInput`, `Type`, `Value` |
| `Watch` | `Application`, `Delete`, `Parent`, `Source` |
| `Watches` | `Add`, `Application`, `Count`, `Delete`, `Item`, `Parent` |
| `Window` | `Activate`, `ActivateNext`, `ActivatePrevious`, `ActiveCell`, `ActiveSheet`, `Application`, `Caption`, `Close`, `Index`, `NewWindow`, `Parent`, `PrintOut`, `PrintPreview`, `RangeSelection`, `SelectedSheets`, `Selection`, `Type`, `Visible`, `WindowNumber` |
| `Windows` | `Application`, `Count`, `Item`, `Parent` |
| `Workbook` | `Activate`, `ActiveSheet`, `Application`, `BuiltinDocumentProperties`, `Calculate`, `Close`, `CodeName`, `CustomViews`, `ExportAsFixedFormat`, `FullName`, `IconSets`, `Name`, `Names`, `NewWindow`, `Parent`, `Path`, `PivotCaches`, `PrintOut`, `PrintPreview`, `Protect`, `ProtectStructure`, `Save`, `SaveAs`, `SaveCopyAs`, `Saved`, `Sheets`, `SlicerCaches`, `Styles`, `Unprotect`, `Windows`, `Worksheets` |
| `Workbooks` | `Add`, `Application`, `Count`, `Item`, `Open`, `Parent` |
| `Worksheet` | `Activate`, `Application`, `AutoFilter`, `AutoFilterMode`, `Cells`, `ChartObjects`, `CodeName`, `Columns`, `Comments`, `CommentsThreaded`, `Copy`, `Delete`, `EnableSelection`, `Evaluate`, `ExportAsFixedFormat`, `FilterMode`, `HPageBreaks`, `Hyperlinks`, `Index`, `ListObjects`, `Move`, `Name`, `Names`, `Outline`, `PageSetup`, `Parent`, `Paste`, `PivotTables`, `PrintOut`, `PrintPreview`, `Protect`, `ProtectContents`, `ProtectDrawingObjects`, `ProtectScenarios`, `Protection`, `ProtectionMode`, `Range`, `ResetAllPageBreaks`, `Rows`, `Scenarios`, `Select`, `Shapes`, `ShowAllData`, `Sort`, `Unprotect`, `UsedRange`, `VPageBreaks`, `Visible` |
| `WorksheetFunction` | `Application`, `Parent` |
| `Worksheets` | `Add`, `Application`, `Count`, `Item`, `Parent`, `PrintOut`, `PrintPreview` |

</details>

<!-- END VBA OBJECT MEMBERS -->

M7A revision — 2026-11-04: initial native worksheet metadata decoder and explicit
remaining native form/control gaps; no interactive control or edited-save claim.

M7A.1/M7A.2 revision — 2026-11-04: native package UserForms attach inertly via
the shared decoder; explicit consent precedes execution. [Profile and contracts](../../../tiny-vba/docs/m7a-control-contracts.md).

M7A.3 revision — 2026-11-04: native worksheet controls supersede the earlier
metadata-only slice; reference catalogue regenerated against the existing shared
87a862a9 pin. Fresh Office comparisons remain NOT_RUN.

M7A.4 increment — 2026-11-04: the pinned shared runtime adds explicit instances,
selected boolean/group/input/property semantics, Controls/UserForms lookup and
separate window/code lifetime. See [the shared scope and remaining work](../../../tiny-vba/docs/m7a4-native-runtime.md).
This is not full M7A.4 or native edited-save acceptance.

M7A.4 runtime follow-up — 2026-11-04: native text selection, multiline/password
and input behavior plus selected SpinButton/ScrollBar properties/events now use
the shared canonical runtime and trusted input path. Further native control,
property, binding and platform work remains in the shared implementation plan.

M7A.4 ListBox follow-up — 2026-11-04: selected native unbound single-selection
ListBox rows, indexed List, AddItem/RemoveItem/Clear, Value/Text/columns and
trusted pointer/keyboard input use the shared runtime. Full list APIs, multiple
selection/bindings and ComboBox remain incomplete; see the shared runtime record.

M7A.4 ComboBox follow-up — 2026-11-04: selected native unbound editable/drop-list
ComboBox text/selection/matching, rows/columns and drop input/events use the shared
canonical runtime. Locale limits, further list/binding/resource and lifecycle
work remain in the shared implementation plan.

M7A.4 list array follow-up — 2026-11-04: native ListBox/ComboBox support bounded
2D List loading, transposed Column loading and indexed/current-row Column access.
Source arrays remain independent; validation precedes row/selection replacement
and ordinary Change continuations. See [integration evidence](integration.md).

M7A.4 readback/metadata follow-up — 2026-11-04: native ListBox/ComboBox expose
detached 2D List/Column snapshots for the current ten-column unbound packet.
Native Tag/ControlTipText/HelpContextID use ordinary per-instance state; actual
control-tip hover is read-only and theme-injected. See [integration](integration.md).

M7A.4 native Frame follow-up (2026-11-05): the shared runtime now supports
selected nested native Frame storages, immediate child Controls/Parent identity,
scoped tab/option groups, nested ActiveControl and cancellable ancestor focus.
Shared widgets use live relative geometry, clipping and accessible group ancestry.
The package regression exercises typed Frame access and trusted child text input,
then verifies exact original project resources after save/reopen. This does not
implement native edited-source/resource saving. MultiPage/TabStrip/Image/RefEdit,
further properties/bindings/platform/lifecycle semantics and .tvform removal
remain open; fresh Office observations remain NOT_RUN under M8.5.

M7A.4 native TabStrip follow-up — 2026-11-05: selected native records now have
separate Tab/Tabs identity, indexed/named lookup, enumeration, Value/SelectedItem,
metadata and authorized pointer/arrow input in the shared runtime. Nearby controls
stay on their original form/Frame. Dynamic tab mutation, writable Index/Name,
further presentation/events and the remaining M7A.4 scope remain open. Native
edited saving and .tvform removal remain separate milestones.

M7A.4 native MultiPage/Page follow-up — 2026-11-05: shared revision
`18e412caa8d9dfd1a1ba984ea49c42bc490ef8da` adds bounded native f/o/x joins,
canonical Page/Pages identity and child Controls/Parent, indexed/named lookup,
enumeration, Value/SelectedItem and selected-page input/focus/rendering. Internal
tabs remain private. Physical selection respects dirty-field cancellation and
Stop/resume. TabStrip/MultiPage Click now carries its Long index and selecting
the current page/tab raises no selection event. Native source/resources remain
inert and separate; original project bytes survive package save/reopen. Dynamic
Pages/Tabs mutation, further page/Frame/property/resource/event/lifecycle cases,
Image/RefEdit, bindings and applicable platform controls remain M7A.4 requirements.
Native edited saving/.tvform removal remain M7A.5/.6; no full milestone sign-off
or fresh Office observations are claimed.

M7A.4 native Image/picture follow-up — 2026-11-05: shared revision
`4455df4082430d3d0d56412df73defd92c26c1df` adds bounded inert BMP/GIF/JPEG/ICO
resources, ordinary Image/stdole picture identity and typed assignment/clearing,
clip/stretch/zoom, alignment/tiling and selected AutoSize/border/effect state.
Image clicks preserve dirty text focus and ordinary callback/Stop continuations.
Canonical state stays in VM instances, with immutable shared pixel assets and
bounded derivative texture caches. Original source/resources remain separate.
Vector/platform picture services and the full image event/presentation matrix,
RefEdit, bindings, dynamic collections, platform controls and remaining lifecycle/
property semantics stay in M7A.4. Native edited saving/.tvform removal remain
M7A.5/.6; fresh Office observations remain NOT_RUN under M8.5.

M7A.4 native RefEdit follow-up — 2026-11-05: the shared implementation adds
bounded inert class-table/CLSID resolution, ordinary RefEdit reference metadata,
canonical text/drop state and selected focus/validation/Change/DropButtonClick
continuations. Original source/resources remain separate. The Excel host owns
worksheet range selection; no DLL/OCX/COM activation is introduced. Broader
RefEdit/platform/property/event semantics, bindings, dynamic collections and the
remaining M7A.4 ledger stay open. Native edited saving/removal stay M7A.5/.6;
fresh Office observations remain NOT_RUN under M8.5. Integration checks PASS;
see the RefEdit validation record in [integration](integration.md).

M7A.4 native Tabs/Pages mutation — 2026-11-05 (shared revision
`422c5c2338a05276e353b27018a531bf23231af4`): native collections support
Add/Remove/Clear, writable Integer Index and runtime-created member Name.
Reordering preserves aliases and selected identity; collection changes commit
before ordinary Change/Stop continuations. Added empty Pages and reordered
native children use the shared renderer; removed Page children reject widget
input. Original project/source/resource bytes remain separate. Fresh Office
comparison remains NOT_RUN under M8.5; Controls.Add/Remove, multiple selection,
bindings and the remaining M7A.4 semantic ledger stay open. Native edited saving
and .tvform removal remain M7A.5/.6. Integration validation PASS: all 13 native and 13 actual Node/WASM package
cases, full ten-step CI, strict optional VBA/UI Clippy, three form widget cases
and the optional VBA web check. The new case checks both Tabs and Pages,
authorization, retained aliases and unchanged XLSM VBA project bytes.


M7A.4 runtime controls/selection/bindings — 2026-11-05: selected native runtime
Controls.Add/Remove/Clear, ListBox Multi/Extended selection and canonical native
TypeName are implemented. ControlSource links one worksheet cell; RowSource loads
bounded contiguous worksheet ranges into ListBox/ComboBox. Imported binding
metadata attaches during normal trusted form load before Initialize. Programmatic
and authorized widget changes use existing cell commands; worksheet changes queue
ordinary form callbacks. Source inspection/rendering remain inert. Original XLSM
VBA resources remain exact; native edited form persistence remains separate.


Runtime type inspection (2026-11-06): TypeName now reports native names from the
existing workbook object store through Object/Variant/interface aliases without
reading default properties. TypeName/VarType/IsArray accept varname named arguments.
The existing authorization and execution/memory limits remain unchanged.

VarType default values (2026-11-06, ADR 0067/IR 13): compatible native control
and project defaults use ordinary frames; Range.Value and Application.Name use
the existing authorized host call path. Aliases retain identity, getter errors
propagate, and Stop/host waits do not repeat evaluation. Typed Date/Currency and
array element tags are preserved. TypeName/IsObject stay inert reference queries.
Document and undo invariance are covered by authored host regressions; fresh
Office comparisons remain M8.5 NOT_RUN. Broader M7A.4 remains open.

Pinned shared revision: `5f93e8140b457de6645327b885b263ab6b472e13`.
Validation PASS: full CI (10 steps), 12 native/12 Node-WASM identity cases,
strict optional VBA/UI Clippy, four form UI cases and optional web build.
Explicit empty indexed collection calls retain their missing-argument error;
VarType treats those collections as objects without invoking Item.


As New arrays (2026-11-06, shared ADR 0068/IR 14): fixed/dynamic project-class
and native-form elements construct lazily through ordinary frames. Element
identity, Nothing recreation, Preserve/Erase, private/static/restored storage,
canonical ByRef and Stop/For Each continuations preserve existing budgets.
A new authored XLSM test checks independent window callbacks, trust revocation
and exact original project resources. Pinned validation PASS at
`d3f00127a47f08996bd89968ebb7ec019ed13c8b`: full CI (10 steps), strict optional
VBA/UI Clippy, four form UI cases, optional web build and all 20 native/20
Node-WASM package cases. Broader M7A.4 and fresh Office comparisons remain open.


UserForms without form components — 2026-11-06 (shared ADR 0069):
the existing built-in collection supports zero Count, empty iteration, aliases
and runtime type names, Item bounds errors and error 424 for Add of a missing
form. Normal semantic shadowing and qualified VBA.UserForms remain available.
No control library or form is loaded by the no-component provider. Authored host
cases check authorization, Stop/resume and unchanged document/undo state.
Pinned validation PASS at `0e7a3fea4f306a1aac8207e00b0a2deb206f6737`: full CI
(10 steps), strict optional VBA/UI Clippy, four form UI cases, optional web and
all 13 native/13 Node-WASM identity cases. Broader M7A.4 remains open.

M7A.4 native keyboard follow-up — 2026-11-06: shared revision
`9fde04ce9ed8eab15b7f0609160617c774c2fc1b` adds selected typed
KeyDown/KeyPress/KeyUp, mutable ByVal MSForms.ReturnInteger, Shift flags and
static/dynamic WithEvents delivery. The shared UI waits before applying default
input, preserves Stop/Continue and discards stale Reset input. Zero cancels;
KeyPress replacement uses existing injected ANSI services. Unicode without a
connected KeyPress handler retains the existing text path. All 21 native and 21 Node-WASM keyboard/package cases pass, alongside full
CI (10 steps), strict optional VBA/UI Clippy, four form UI cases and optional
web compilation. Full keyboard/platform/IME and M7A.4 remain partial;
see [ADR 0070](../../../tiny-vba/docs/decisions/0070-native-keyboard-events.md).

M7A.4 keyboard-focus update — 2026-11-06: shared revision
`38578095ba1fc1ae3695f92a6c2e89cdb29d42f6` adds selected root focus and
Tab/SetFocus/Space ordering, stopped character/release continuation, native
function-key ownership and web text/key normalization. Root handlers require
eligible form focus; setting TabStop=False alone does not grant it. Full shared
gates, nine runtime keyboard cases, eleven new UI cases, the existing 24 form UI
cases and Chromium pass. Full CI (10 steps), strict optional VBA/UI Clippy, four form UI cases, web and
all 22 native/22 Node-WASM package cases pass, including root eligibility,
Stop/Continue, authority and original VBA resources.
See [ADR 0071](../../../tiny-vba/docs/decisions/0071-native-keyboard-focus-and-default-order.md).

M7A.4 accelerator update (2026-11-06): shared ADR 0072 at `7d43f082de4b7827da1e1d93c257ebeb50eca4ae`
adds persisted native accelerators, canonical Accelerator properties and selected
Label/button/boolean/Tab/Page Alt actions. Focus validation, Stop/Continue and
existing authorization apply. Caption underlines use the host theme and fonts.
Duplicate routing explicitly refuses; non-ASCII/platform ordering, mouse,
lifecycle and other checklist rows remain open. M7A.4 remains incomplete; fresh
Office observations remain M8.5 NOT_RUN. Full CI (10 steps), strict optional VBA/UI Clippy, four form UI cases, web and
all 23 native/23 Node-WASM package cases pass. The accelerator case preserves
authorization, one-shot debugger continuation, workbook effects and original
VBA resources.

M7A.4 mouse update (2026-11-06): shared ADR 0073 adds typed native
MouseDown/MouseMove/MouseUp/DblClick callbacks, indexed arguments, mutable cancel
and selected simple-control capture/defaults. Stop/Continue, outside release and
Reset preserve input identity and source. Event signatures omit unused class-field
metadata while preserving typed arguments, executable method fields and all limits.
Text/list/numeric/indexed-tab pointer defaults, lifecycle and other checklist rows
remain open; M7A.4 is incomplete and Office observations remain M8.5 NOT_RUN.
Pinned shared revision `c9f829824c56397c101ac9228c89db289b788458` passes full
CI (10 steps), strict optional VBA/UI Clippy, four form UI tests, web and all
24 native/24 Node-WASM package cases. Indexed MultiPage mouse input, stopped
callbacks, authority, workbook effects and exact original resources are verified.

Mouse command authorization is explicit under shared ADR 0073: form commands
require host preparation unless they are known inert design operations. The final
pin passes the same full CI/optional UI/web and 24 native/24 Node-WASM gates.

M7A.4 TextBox pointer update (2026-11-06): shared ADR 0074 adds selected
caret, captured drag, double-click cancellation/word selection and Shift extension
after ordinary mouse callbacks. Native AutoWordSelect is decoded with its inverse
file bit; new dynamic text controls retain their separate default. Canonical
UTF-16/CRLF, focus redirects, host authorization, Stop/Continue and Reset/source
identity remain enforced. Modal bounds stay stable while input waits. Editable
ComboBox/RefEdit, list/numeric/tab pointer defaults, further platform/selection and
lifecycle cases remain open; M7A.4 is incomplete and fresh Office comparisons
remain M8.5 NOT_RUN. Validation receipts follow below.

Accepted shared pin `3e8676cc754f59e06df64d2677c1f64af0216432`: full CI (10 steps), strict optional VBA/UI Clippy,
25 native and 25 actual Node-WASM package cases, four form UI cases and web checks
pass. The new selection case preserves authority, stopped continuation and exact
original workbook/VBA resources. M7A.4 remains incomplete.

M7A.4 composite-pointer update (2026-11-06): shared ADR 0075 adds selected
ComboBox/RefEdit text/drop/picker, list/popup row, numeric step/repeat/drag and
indexed tab/page-header defaults after ordinary callbacks. Retained modal display
metadata preserves window presentation while callbacks run, without authorizing
input. Tabs retain identity through reorder and discard removed targets. Existing
host consent, snapshots, ordinary VBA state and resource preservation remain in
force; ComboBox AutoWordSelect preserves its inverse native bit. M7A.4 and the
full runtime checklist remain incomplete. Fresh Office observations are M8.5
NOT_RUN; native edited persistence/removal remain separate.
Pinned shared product `f027c5d14f0a57c6a0ebfa0f1cc2fb393cb9c2b7` passes full CI (10 steps),
strict optional VBA/UI Clippy, four form UI cases, the VBA web build and all
26 native/26 actual Node-WASM package cases. The new modal pointer regression
preserves stopped authority, one-shot workbook effects and original resources.

M7A.4 nested-input update (2026-11-06): shared ADR 0076 services child mouse
callbacks while a parent mouse/key callback is parked. Raw input scopes share the
existing 64-event queue; held releases retain their original capture, modal areas
follow Show order, and open ComboBox popups stay above their owning form. Reset,
parent unload and owner-local keyboard barriers suppress stale/replayed defaults.
Canonical VBA state and existing host authority remain unchanged. Selected L02
cases are covered; M7A.4 and the full lifecycle/control checklist remain partial.
Fresh Office observations remain M8.5 NOT_RUN.
Pinned shared product `c99996351a80a234b23f2e81b54e210dae0bd530` passes full CI (10 steps),
strict optional VBA/UI Clippy, four form UI cases, VBA web and all 27 native/27
actual Node-WASM package cases. The nested callback regression preserves modal
ownership, separate child/parent receipts, one-shot workbook effects and original
resources.

ADR 0077 adds selected native Show/Hide/modeless Activate/Deactivate ordering,
unload retirement without Deactivate, reentrant lifecycle guards and cross-window
keyboard continuations. Inert inspection and executable window-focus commands
retain existing trust. Automatic document-canvas input barriers, broader binding/
collection/platform ordering and full M7A.4 acceptance remain open.
