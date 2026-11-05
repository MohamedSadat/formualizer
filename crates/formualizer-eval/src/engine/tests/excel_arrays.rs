//! How ranges and arrays behave in formulas, as Excel's stored values show.

use crate::engine::named_range::{NameScope, NamedDefinition};
use crate::engine::{Engine, EvalConfig};
use crate::test_workbook::TestWorkbook;
use formualizer_common::{ExcelErrorKind, LiteralValue};
use formualizer_parse::parser::parse;

fn engine() -> Engine<TestWorkbook> {
    Engine::new(TestWorkbook::new(), EvalConfig::default())
}

fn text(value: &str) -> LiteralValue {
    LiteralValue::Text(value.to_string())
}

fn number(value: f64) -> LiteralValue {
    LiteralValue::Number(value)
}

/// `A1:A3` holds x, y, x and `B1:B3` holds 1, -2, 3.
fn with_data() -> Engine<TestWorkbook> {
    let mut e = engine();
    for (row, (a, b)) in [("x", 1.0), ("y", -2.0), ("x", 3.0)].iter().enumerate() {
        let row = row as u32 + 1;
        e.set_cell_value("Sheet1", row, 1, text(a)).unwrap();
        e.set_cell_value("Sheet1", row, 2, number(*b)).unwrap();
    }
    e
}

fn evaluated(mut e: Engine<TestWorkbook>, formula: &str) -> LiteralValue {
    e.set_cell_formula("Sheet1", 20, 20, parse(formula).unwrap())
        .unwrap();
    e.evaluate_all().unwrap();
    e.get_cell_value("Sheet1", 20, 20).unwrap()
}

#[test]
fn a_single_value_parameter_runs_once_for_each_value_of_an_array() {
    for (formula, expected) in [
        ("=SUM(SIGN({1,-2,3}))", 1.0),
        ("=SUM(ABS(B1:B3))", 6.0),
        ("=SUM(LEN(A1:A3))", 3.0),
        ("=SUM(MATCH(A1:A3,A1:A3,0))", 4.0),
        ("=SUM(ROUND(B1:B3/2,0))", 2.0),
        ("=SUMPRODUCT(ISNUMBER(MATCH(A1:A3,{\"x\"},0))*1)", 2.0),
    ] {
        assert_eq!(
            evaluated(with_data(), formula),
            number(expected),
            "{formula}"
        );
    }
}

#[test]
fn arrays_of_different_sizes_run_to_the_larger() {
    assert_eq!(
        evaluated(engine(), "=SUM(LEN(LEFT({\"ab\";\"cd\"},{1,2})))"),
        number(6.0),
        "a column and a row go with each other"
    );
    assert_eq!(
        evaluated(engine(), "=ISNA(INDEX(POWER({1;2;3},{1;2}),3))"),
        LiteralValue::Boolean(true),
        "past the end of the shorter array there is nothing to read"
    );
}

#[test]
fn an_error_before_any_array_stops_the_call() {
    assert_eq!(
        evaluated(engine(), "=ERROR.TYPE(ROUND(1/0,LEN({\"a\",\"b\"})))"),
        number(2.0)
    );
}

#[test]
fn row_and_column_give_every_row_and_column_of_a_range() {
    assert_eq!(evaluated(with_data(), "=SUM(ROW(A1:A3))"), number(6.0));
    assert_eq!(evaluated(with_data(), "=SUM(COLUMN(A1:C9))"), number(6.0));
    assert_eq!(
        evaluated(with_data(), "=ROW(B2)"),
        LiteralValue::Number(2.0)
    );
}

#[test]
fn frequency_skips_what_isnt_a_number_in_an_array_but_refuses_a_single_one() {
    assert_eq!(
        evaluated(engine(), "=SUM(FREQUENCY({1,FALSE,\"a\",3},{2,4}))"),
        number(2.0)
    );
    let LiteralValue::Error(error) = evaluated(engine(), "=SUM(FREQUENCY(FALSE,{2,4}))") else {
        panic!("a single value that isn't a number is an error");
    };
    assert_eq!(error.kind, ExcelErrorKind::Value);
}

#[test]
fn index_reads_a_table_column() {
    let mut e = engine();
    e.set_cell_value("Sheet1", 1, 1, text("V")).unwrap();
    e.set_cell_value("Sheet1", 2, 1, text("p")).unwrap();
    e.set_cell_value("Sheet1", 3, 1, text("q")).unwrap();
    let id = e.sheet_id("Sheet1").unwrap();
    let start =
        crate::reference::CellRef::new(id, crate::reference::Coord::from_excel(1, 1, true, true));
    let end =
        crate::reference::CellRef::new(id, crate::reference::Coord::from_excel(3, 1, true, true));
    e.define_table(
        "T",
        crate::reference::RangeRef::new(start, end),
        true,
        vec!["V".to_string()],
        false,
    )
    .unwrap();
    assert_eq!(evaluated(e, "=INDEX(T[V],2)"), text("q"));
}

#[test]
fn at_keeps_a_blank_blank_and_a_single_cell_range_reads_from_anywhere() {
    assert_eq!(evaluated(engine(), "=@Z:Z&\"/\""), text("/"));
    let mut e = with_data();
    e.set_cell_value("Sheet1", 20, 2, text("level")).unwrap();
    assert_eq!(evaluated(e, "=@B:B"), text("level"));
    assert_eq!(evaluated(with_data(), "=@B3:B3"), number(3.0));
}

#[test]
fn the_intersection_operator_reads_what_two_references_share() {
    assert_eq!(evaluated(with_data(), "=SUM(A1:B2 B2:B3)"), number(-2.0));
    assert_eq!(
        evaluated(with_data(), "=Sheet1!B:B Sheet1!3:3"),
        number(3.0)
    );
    let LiteralValue::Error(error) = evaluated(with_data(), "=SUM(A1:A2 B1:B2)") else {
        panic!("two references that share nothing are #NULL!");
    };
    assert_eq!(error.kind, ExcelErrorKind::Null);
}

#[test]
fn the_union_operator_reads_all_its_references() {
    assert_eq!(evaluated(with_data(), "=SUM((B1,B2:B3))"), number(2.0));
    assert_eq!(evaluated(with_data(), "=COUNT((A1:A3,B1,B3))"), number(2.0));
    assert_eq!(
        evaluated(with_data(), "=RANK(B1,(B1,B2,B3),1)"),
        number(2.0)
    );
}

#[test]
fn single_is_how_a_file_writes_at() {
    let mut e = with_data();
    e.set_cell_value("Sheet1", 20, 2, number(7.0)).unwrap();
    assert_eq!(evaluated(e, "=_xlfn.SINGLE(B:B)+1"), number(8.0));
    assert_eq!(evaluated(engine(), "=SINGLE({5,6})"), number(5.0));
}

fn define(e: &mut Engine<TestWorkbook>, name: &str, formula: &str) {
    let ast = parse(formula).unwrap();
    e.define_name(
        name,
        NamedDefinition::Formula {
            ast,
            dependencies: Vec::new(),
            range_deps: Vec::new(),
        },
        NameScope::Workbook,
    )
    .unwrap();
}

#[test]
fn a_name_that_stands_for_a_reference_or_an_array_reads_all_of_it() {
    let mut e = with_data();
    define(
        &mut e,
        "Dynamic",
        "=OFFSET(Sheet1!$B$1,0,0,COUNT(Sheet1!$B:$B),1)",
    );
    assert_eq!(evaluated(e, "=SUM(Dynamic)"), number(2.0));
    let mut e = engine();
    define(&mut e, "Constants", "={\"A\",\"B\"}");
    assert_eq!(evaluated(e, "=INDEX(Constants,1,2)"), text("B"));
}

#[test]
fn a_name_can_be_read_through_its_sheet() {
    let mut e = with_data();
    e.add_sheet("Other").unwrap();
    let id = e.sheet_id("Sheet1").unwrap();
    e.define_name(
        "Local",
        NamedDefinition::Literal(number(5.0)),
        NameScope::Sheet(id),
    )
    .unwrap();
    e.set_cell_formula("Other", 1, 1, parse("=Sheet1!Local+1").unwrap())
        .unwrap();
    e.evaluate_all().unwrap();
    assert_eq!(e.get_cell_value("Other", 1, 1).unwrap(), number(6.0));
}

#[test]
fn a_named_formula_works_out_its_value_after_stored_values_are_adopted() {
    let mut e = with_data();
    define(&mut e, "Total", "=SUM(Sheet1!$B$1:$B$3)");
    e.set_cell_formula("Sheet1", 20, 20, parse("=Total+1").unwrap())
        .unwrap();
    e.adopt_stored_values().unwrap();
    e.mark_all_formulas_dirty();
    e.evaluate_all().unwrap();
    assert_eq!(e.get_cell_value("Sheet1", 20, 20).unwrap(), number(3.0));
}

#[test]
fn a_spill_shows_over_an_edit_that_emptied_the_cell() {
    let mut e = with_data();
    e.set_cell_value("Sheet1", 11, 20, text("saved")).unwrap();
    e.set_cell_value("Sheet1", 11, 20, LiteralValue::Empty)
        .unwrap();
    e.set_cell_formula("Sheet1", 10, 20, parse("=B1:B3").unwrap())
        .unwrap();
    e.evaluate_all().unwrap();
    assert_eq!(e.get_cell_value("Sheet1", 11, 20).unwrap(), number(-2.0));
}

#[test]
fn iserror_looks_at_what_cant_be_evaluated() {
    assert_eq!(
        evaluated(engine(), "=ISERROR(VLOOKUP(\"x\",nosuchname,2,FALSE))"),
        LiteralValue::Boolean(true)
    );
    assert_eq!(
        evaluated(engine(), "=IF(ISERROR(NoSheet!A1),\"caught\",\"no\")"),
        text("caught")
    );
    assert_eq!(
        evaluated(engine(), "=ISNA(nosuchname)"),
        LiteralValue::Boolean(false)
    );
}

#[test]
fn isformula_and_formulatext_see_the_formula() {
    let mut e = engine();
    e.set_cell_value("Sheet1", 1, 1, number(1.0)).unwrap();
    e.set_cell_formula("Sheet1", 2, 1, parse("=_xlfn._xlws.SORT(A1)").unwrap())
        .unwrap();
    e.set_cell_formula(
        "Sheet1",
        3,
        1,
        parse("=SUM(_xlfn.ANCHORARRAY(A2))").unwrap(),
    )
    .unwrap();
    e.set_cell_formula("Sheet1", 1, 2, parse("=ISFORMULA(A1)").unwrap())
        .unwrap();
    e.set_cell_formula("Sheet1", 2, 2, parse("=ISFORMULA(A2)").unwrap())
        .unwrap();
    e.set_cell_formula("Sheet1", 2, 3, parse("=FORMULATEXT(A2)").unwrap())
        .unwrap();
    e.set_cell_formula("Sheet1", 3, 3, parse("=FORMULATEXT(A3)").unwrap())
        .unwrap();
    e.evaluate_all().unwrap();
    assert_eq!(
        e.get_cell_value("Sheet1", 1, 2).unwrap(),
        LiteralValue::Boolean(false)
    );
    assert_eq!(
        e.get_cell_value("Sheet1", 2, 2).unwrap(),
        LiteralValue::Boolean(true)
    );
    let LiteralValue::Text(sort) = e.get_cell_value("Sheet1", 2, 3).unwrap() else {
        panic!("FORMULATEXT gives text");
    };
    assert!(sort.eq_ignore_ascii_case("=SORT(A1)"), "{sort}");
    let LiteralValue::Text(sum) = e.get_cell_value("Sheet1", 3, 3).unwrap() else {
        panic!("FORMULATEXT gives text");
    };
    assert!(sum.eq_ignore_ascii_case("=SUM(A2#)"), "{sum}");
}

#[test]
fn cell_gives_the_sheet_after_the_bracket() {
    assert_eq!(
        evaluated(
            with_data(),
            "=MID(CELL(\"filename\",A1),FIND(\"]\",CELL(\"filename\",A1))+1,99)"
        ),
        text("Sheet1")
    );
}

#[test]
fn lookup_finds_the_last_match_past_errors() {
    assert_eq!(
        evaluated(with_data(), "=LOOKUP(2,1/(A1:A3=\"x\"),B1:B3)"),
        number(3.0)
    );
    let LiteralValue::Error(error) = evaluated(with_data(), "=LOOKUP(Z9,A1:A3)") else {
        panic!("a blank is 0, and there are no numbers here");
    };
    assert_eq!(error.kind, ExcelErrorKind::Na);
}

#[test]
fn sin_and_cos_reduce_by_pi_as_the_processor_has_it() {
    let LiteralValue::Number(sin) = evaluated(engine(), "=SIN(PI())") else {
        panic!("a number");
    };
    assert!((sin - 1.22514845490862e-16).abs() < 1e-28, "{sin}");
    let LiteralValue::Number(cos) = evaluated(engine(), "=COS(PI()/2)") else {
        panic!("a number");
    };
    assert!((cos - 6.1257422745431e-17).abs() < 1e-28, "{cos}");
    assert_eq!(evaluated(engine(), "=COS(0)"), number(1.0));
}

#[test]
fn sumif_reads_an_array_of_indirect_references_one_by_one() {
    let mut e = engine();
    e.add_sheet("One").unwrap();
    e.add_sheet("Two").unwrap();
    for (sheet, value) in [("One", 10.0), ("Two", 20.0)] {
        e.set_cell_value(sheet, 1, 1, text("key")).unwrap();
        e.set_cell_value(sheet, 1, 2, number(value)).unwrap();
    }
    assert_eq!(
        evaluated(
            e,
            "=SUMPRODUCT(SUMIF(INDIRECT({\"One\";\"Two\"}&\"!A:A\"),\"key\",INDIRECT({\"One\";\"Two\"}&\"!B:B\")))"
        ),
        number(30.0)
    );
}

#[test]
fn a_function_given_by_name_reads_as_its_name() {
    let mut e = engine();
    for (row, (group, value)) in [("b", 1.0), ("a", 2.0), ("b", 3.0)].iter().enumerate() {
        let row = row as u32 + 1;
        e.set_cell_value("Sheet1", row, 1, text(group)).unwrap();
        e.set_cell_value("Sheet1", row, 2, number(*value)).unwrap();
    }
    assert_eq!(
        evaluated(e, "=INDEX(GROUPBY(A1:A3,B1:B3,_xleta.SUM,0),1,2)"),
        number(2.0),
        "sorted by group, with a in the first row"
    );
}

#[test]
fn renaming_a_sheet_to_other_letters_changes_its_spelling() {
    let mut e = engine();
    let id = e.sheet_id("Sheet1").unwrap();
    e.rename_sheet(id, "sheet1").unwrap();
    assert_eq!(e.sheet_name(id), "sheet1");
}
