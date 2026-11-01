use crate::engine::{Engine, EvalConfig};
use crate::reference::{CellRef, Coord, RangeRef};
use crate::test_workbook::TestWorkbook;
use formualizer_common::LiteralValue;
use formualizer_parse::parser::parse;

fn number(engine: &mut Engine<TestWorkbook>, row: u32, col: u32, value: f64) {
    engine
        .set_cell_value("Sheet1", row, col, LiteralValue::Number(value))
        .unwrap();
}

fn formula(engine: &mut Engine<TestWorkbook>, row: u32, col: u32, text: &str) {
    engine
        .set_cell_formula("Sheet1", row, col, parse(text).unwrap())
        .unwrap();
}

fn table(engine: &mut Engine<TestWorkbook>, rows: u32, cols: u32, headers: &[&str], totals: bool) {
    let sheet = engine.sheet_id("Sheet1").unwrap();
    let start = CellRef::new(sheet, Coord::from_excel(1, 1, true, true));
    let end = CellRef::new(sheet, Coord::from_excel(rows, cols, true, true));
    engine
        .define_table(
            "Sales",
            RangeRef::new(start, end),
            true,
            headers.iter().map(|h| h.to_string()).collect(),
            totals,
        )
        .unwrap();
}

fn new_engine() -> Engine<TestWorkbook> {
    let mut engine = Engine::new(TestWorkbook::new(), EvalConfig::default());
    engine.add_sheet("Sheet1").unwrap();
    engine
}

#[test]
fn a_totals_row_reading_its_own_table_is_not_a_cycle() {
    let mut engine = new_engine();
    number(&mut engine, 2, 2, 10.0);
    number(&mut engine, 3, 2, 20.0);
    table(&mut engine, 4, 2, &["Item", "Cost"], true);
    formula(&mut engine, 4, 2, "=SUBTOTAL(109,Sales[Cost])");
    engine.evaluate_all().unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", 4, 2),
        Some(LiteralValue::Number(30.0))
    );

    number(&mut engine, 3, 2, 25.0);
    engine.evaluate_all().unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", 4, 2),
        Some(LiteralValue::Number(35.0))
    );
}

#[test]
fn a_reader_of_the_totals_row_follows_it() {
    let mut engine = new_engine();
    number(&mut engine, 2, 2, 10.0);
    number(&mut engine, 3, 2, 20.0);
    table(&mut engine, 4, 2, &["Item", "Cost"], true);
    formula(&mut engine, 4, 2, "=SUBTOTAL(109,Sales[Cost])");
    formula(&mut engine, 1, 4, "=Sales[[#Totals],[Cost]]*2");
    engine.evaluate_all().unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", 1, 4),
        Some(LiteralValue::Number(60.0))
    );

    number(&mut engine, 2, 2, 11.0);
    engine.evaluate_all().unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", 1, 4),
        Some(LiteralValue::Number(62.0))
    );
}

#[test]
fn a_named_this_row_reference_reads_its_row() {
    let mut engine = new_engine();
    number(&mut engine, 2, 1, 1.0);
    number(&mut engine, 2, 2, 2.0);
    number(&mut engine, 3, 1, 3.0);
    number(&mut engine, 3, 2, 4.0);
    table(&mut engine, 3, 4, &["Q1", "Q2", "Sum", "Twice"], false);
    for row in 2..=3 {
        formula(&mut engine, row, 3, "=SUM(Sales[[#This Row],[Q1]:[Q2]])");
        formula(&mut engine, row, 4, "=Sales[[#This Row],[Sum]]*2");
    }
    engine.evaluate_all().unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", 2, 3),
        Some(LiteralValue::Number(3.0))
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", 3, 4),
        Some(LiteralValue::Number(14.0))
    );

    number(&mut engine, 3, 1, 10.0);
    engine.evaluate_all().unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", 3, 4),
        Some(LiteralValue::Number(28.0))
    );
}

#[test]
fn ceiling_and_floor_math_take_a_mode() {
    let mut engine = new_engine();
    formula(&mut engine, 1, 1, "=CEILING.MATH(-2.5,1,1)");
    formula(&mut engine, 2, 1, "=CEILING.MATH(-2.5,1,0)");
    formula(&mut engine, 3, 1, "=FLOOR.MATH(-2.5,1,1)");
    formula(&mut engine, 4, 1, "=FLOOR.MATH(-2.5,1,0)");
    engine.evaluate_all().unwrap();
    let value = |row| engine.get_cell_value("Sheet1", row, 1);
    assert_eq!(value(1), Some(LiteralValue::Number(-3.0)));
    assert_eq!(value(2), Some(LiteralValue::Number(-2.0)));
    assert_eq!(value(3), Some(LiteralValue::Number(-2.0)));
    assert_eq!(value(4), Some(LiteralValue::Number(-3.0)));
}

#[test]
fn a_named_range_shrinks_when_part_of_it_is_deleted() {
    use crate::engine::named_range::{NameScope, NamedDefinition};
    let mut engine = new_engine();
    for row in 1..=6 {
        number(&mut engine, row, 1, f64::from(row));
    }
    let sheet = engine.sheet_id("Sheet1").unwrap();
    let start = CellRef::new(sheet, Coord::from_excel(2, 1, true, true));
    let end = CellRef::new(sheet, Coord::from_excel(4, 1, true, true));
    engine
        .define_name(
            "Block",
            NamedDefinition::Range(RangeRef::new(start, end)),
            NameScope::Workbook,
        )
        .unwrap();
    formula(&mut engine, 1, 3, "=SUM(Block)");
    engine.evaluate_all().unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", 1, 3),
        Some(LiteralValue::Number(9.0))
    );

    // Deleting the range's last row leaves rows 2 and 3.
    engine.delete_rows("Sheet1", 4, 1).unwrap();
    engine.evaluate_all().unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", 1, 3),
        Some(LiteralValue::Number(5.0))
    );

    // Deleting all of what's left gives #REF!.
    engine.delete_rows("Sheet1", 2, 2).unwrap();
    engine.evaluate_all().unwrap();
    assert!(matches!(
        engine.get_cell_value("Sheet1", 1, 3),
        Some(LiteralValue::Error(_))
    ));
}

#[test]
fn rows_and_columns_pass_on_a_reference_error() {
    let mut engine = new_engine();
    formula(&mut engine, 1, 1, "=ROWS(#REF!)");
    formula(&mut engine, 2, 1, "=COLUMNS(#REF!)");
    engine.evaluate_all().unwrap();
    for row in 1..=2 {
        assert!(
            matches!(
                engine.get_cell_value("Sheet1", row, 1),
                Some(LiteralValue::Error(_))
            ),
            "row {row}"
        );
    }
}
