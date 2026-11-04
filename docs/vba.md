# VBA in GridCraft

Updated 2026-11-04: [M7 package acceptance](../../tiny-vba/docs/m7-acceptance.md)
adds reviewed original `.xlsm` specimens and real load/compile/run/save/reopen
workflows. The writer preserves the original VBA project and supported related
parts; it does not encode edited IDE source into that binary project. Loading
never executes handlers or grants consent.

GridCraft can run VBA macros and worksheet functions against its own workbooks.
The optional integration includes a source editor, debugger and a subset of the
Excel object model, backed by GridCraft's existing commands and calculator.

Start with these references:

- [Supported functions and features](vba/supported.md)
- [Unsupported functions and features, and why](vba/unsupported.md)
- [How the integration works](vba/integration.md)

## Enable VBA

VBA is disabled in the default application build. From the repository root, run:

```sh
# Embedded editor
cargo run -p gridcraft --features vba

# Separate desktop editor window
cargo run -p gridcraft --features vba-window
```

`vba-window` includes `vba` and uses the same project and execution session.
Browser builds enable `gridcraft-web/vba` and embed the editor in the page.
The pinned tiny-vba dependency supplies the compiler, interpreter and editor;
you do not need a sibling repository checkout.

## Run a macro

1. Open or create a workbook. Enable **Developer** under **File → Options →
   Customize Ribbon → Main Tabs** if the tab is hidden.
2. Choose **Developer → Visual Basic**, or press **Alt+F11**. The editor opens
   without executing imported source.
3. Add a standard module and enter a procedure, for example:

   ```vb
   Public Sub WriteTotal()
       Dim sheet As Worksheet
       Set sheet = ThisWorkbook.Worksheets(1)
       sheet.Range("A1").Value2 = 12
       sheet.Range("A2").Value2 = 30
       sheet.Range("A3").Formula = "=SUM(A1:A2)"
   End Sub
   ```

4. Give explicit VBA permission when prompted, then select the procedure and
   **Run**. Breakpoints, stepping, locals and diagnostics are available in the editor.
5. Use **Save Source to Workbook** to retain standard-module source, then save
   the workbook normally. You can also export modules as plain source files.

**Alt+F8** or **View → Macros → View Macros** opens the macro picker. It lists
public, no-argument Subs. **Edit** and **Step Into** open the selected procedure.
Macro assignments to worksheet shapes and the Quick Access Toolbar are also
available. Design Mode lets you edit assigned shapes without running them.

## Use a worksheet function

Apply a public Function with **Save Source to Workbook**, then use it in a cell:

```vb
Public Function Triple(ByVal number As Double) As Double
    Triple = number * 3
End Function
```

`=Triple(A1)` returns 21 when A1 contains 7. Applied Functions also appear in
**Insert Function → User Defined**. Worksheet functions have a read-only workbook
context and stricter execution limits than macros.

## Opening Excel workbooks

Imported VBA source is inert until you explicitly permit and run it. GridCraft
preserves an existing `vbaProject.bin`, but source edits are saved in GridCraft's
own workbook storage; they do not rewrite Excel's binary VBA project.

Compatibility depends on the functions, members, options and external services a
macro uses. Check the [supported reference](vba/supported.md) and
[known restrictions](vba/unsupported.md) before relying on an existing Excel macro.
