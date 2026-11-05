//! What each parameter of a function does with a range or an array, as Excel has it.
//!
//! A parameter takes a single value, a reference, or an array. Given several values, a
//! single-value parameter runs the function once for each of them, and the result is an array:
//! `LEN(A1:A3)` in an array formula, or `SIGN({1,-2,3})` anywhere. A plain formula in a file reads
//! one cell of a range in those places instead, and that is for whoever loads the file to mark
//! with `@`, using this same table.

use std::collections::HashMap;
use std::sync::OnceLock;

/// What a parameter does with what it is given.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum ParamClass {
    /// A single value, and array mode reaches through it.
    Value,
    /// A single value that stays out of the array mode of what is around it. `VLOOKUP`'s first
    /// parameter is one, so `SUMPRODUCT(VLOOKUP(A1:A9,T,2,0))` needs array entry.
    Scalar,
    /// A reference: a range stays whole, but what an operator works on is values.
    Reference,
    /// `IF`'s branches and the like: what they are given goes straight to `IF`'s own result.
    Branch,
    /// An array, which puts everything under it in array mode.
    Array,
}

#[derive(Clone, Copy)]
enum Params {
    All(ParamClass),
    /// The first parameters, then the class of all after them.
    Listed(&'static [ParamClass], ParamClass),
    /// `COUNTIFS`: range and criteria in turn.
    RangeCriteria,
    /// `SUMIFS`: a range, then range and criteria in turn.
    SumRangeCriteria,
}

fn table() -> &'static HashMap<&'static str, Params> {
    use ParamClass::{Array, Branch, Reference, Scalar, Value};
    static TABLE: OnceLock<HashMap<&'static str, Params>> = OnceLock::new();
    TABLE.get_or_init(|| {
        let mut table = HashMap::new();
        for name in [
            "ABS",
            "ACOS",
            "ACOSH",
            "ACOT",
            "ACOTH",
            "ADDRESS",
            "ARABIC",
            "ASC",
            "ASIN",
            "ASINH",
            "ATAN",
            "ATAN2",
            "ATANH",
            "BAHTTEXT",
            "BASE",
            "BESSELI",
            "BESSELJ",
            "BESSELK",
            "BESSELY",
            "BETA.DIST",
            "BETA.INV",
            "BETADIST",
            "BETAINV",
            "BIN2DEC",
            "BIN2HEX",
            "BIN2OCT",
            "BINOM.DIST",
            "BINOM.DIST.RANGE",
            "BINOM.INV",
            "BINOMDIST",
            "BITAND",
            "BITLSHIFT",
            "BITOR",
            "BITRSHIFT",
            "BITXOR",
            "CEILING",
            "CEILING.MATH",
            "CEILING.PRECISE",
            "CHAR",
            "CHIDIST",
            "CHIINV",
            "CHISQ.DIST",
            "CHISQ.DIST.RT",
            "CHISQ.INV",
            "CHISQ.INV.RT",
            "CLEAN",
            "CODE",
            "COMBIN",
            "COMBINA",
            "COMPLEX",
            "CONCATENATE",
            "CONFIDENCE",
            "CONFIDENCE.NORM",
            "CONFIDENCE.T",
            "CONVERT",
            "COS",
            "COSH",
            "COT",
            "COTH",
            "COUPDAYBS",
            "COUPDAYS",
            "COUPDAYSNC",
            "COUPNCD",
            "COUPNUM",
            "COUPPCD",
            "CRITBINOM",
            "CSC",
            "CSCH",
            "CUMIPMT",
            "CUMPRINC",
            "DATE",
            "DATEDIF",
            "DATEVALUE",
            "DAY",
            "DAYS",
            "DAYS360",
            "DB",
            "DBCS",
            "DDB",
            "DEC2BIN",
            "DEC2HEX",
            "DEC2OCT",
            "DECIMAL",
            "DEGREES",
            "DELTA",
            "DISC",
            "DOLLAR",
            "DOLLARDE",
            "DOLLARFR",
            "DURATION",
            "EDATE",
            "EFFECT",
            "ENCODEURL",
            "EOMONTH",
            "ERF",
            "ERF.PRECISE",
            "ERFC",
            "ERFC.PRECISE",
            "ERROR.TYPE",
            "EVEN",
            "EXACT",
            "EXP",
            "EXPON.DIST",
            "EXPONDIST",
            "F.DIST",
            "F.DIST.RT",
            "F.INV",
            "F.INV.RT",
            "FACT",
            "FACTDOUBLE",
            "FDIST",
            "FILTERXML",
            "FIND",
            "FINDB",
            "FINV",
            "FISHER",
            "FISHERINV",
            "FIXED",
            "FLOOR",
            "FLOOR.MATH",
            "FLOOR.PRECISE",
            "FV",
            "GAMMA",
            "GAMMA.DIST",
            "GAMMA.INV",
            "GAMMADIST",
            "GAMMAINV",
            "GAMMALN",
            "GAMMALN.PRECISE",
            "GAUSS",
            "GESTEP",
            "HEX2BIN",
            "HEX2DEC",
            "HEX2OCT",
            "HOUR",
            "HYPERLINK",
            "HYPGEOM.DIST",
            "HYPGEOMDIST",
            "IMABS",
            "IMAGINARY",
            "IMARGUMENT",
            "IMCONJUGATE",
            "IMCOS",
            "IMCOSH",
            "IMCOT",
            "IMCSC",
            "IMCSCH",
            "IMDIV",
            "IMEXP",
            "IMLN",
            "IMLOG10",
            "IMLOG2",
            "IMPOWER",
            "IMREAL",
            "IMSEC",
            "IMSECH",
            "IMSIN",
            "IMSINH",
            "IMSQRT",
            "IMSUB",
            "IMTAN",
            "INDIRECT",
            "INT",
            "INTRATE",
            "IPMT",
            "ISBLANK",
            "ISERR",
            "ISERROR",
            "ISEVEN",
            "ISLOGICAL",
            "ISNA",
            "ISNONTEXT",
            "ISNUMBER",
            "ISO.CEILING",
            "ISODD",
            "ISOWEEKNUM",
            "ISPMT",
            "ISTEXT",
            "JIS",
            "LEFT",
            "LEFTB",
            "LEN",
            "LENB",
            "LN",
            "LOG",
            "LOG10",
            "LOGINV",
            "LOGNORM.DIST",
            "LOGNORM.INV",
            "LOGNORMDIST",
            "LOWER",
            "MDURATION",
            "MID",
            "MIDB",
            "MINUTE",
            "MOD",
            "MONTH",
            "MROUND",
            "NEGBINOM.DIST",
            "NEGBINOMDIST",
            "NOMINAL",
            "NORM.DIST",
            "NORM.INV",
            "NORM.S.DIST",
            "NORM.S.INV",
            "NORMDIST",
            "NORMINV",
            "NORMSDIST",
            "NORMSINV",
            "NOT",
            "NPER",
            "NUMBERVALUE",
            "OCT2BIN",
            "OCT2DEC",
            "OCT2HEX",
            "ODD",
            "ODDFPRICE",
            "ODDFYIELD",
            "ODDLPRICE",
            "ODDLYIELD",
            "PDURATION",
            "PERMUT",
            "PERMUTATIONA",
            "PHI",
            "PHONETIC",
            "PMT",
            "POISSON",
            "POISSON.DIST",
            "POWER",
            "PPMT",
            "PRICE",
            "PRICEDISC",
            "PRICEMAT",
            "PROPER",
            "PV",
            "QUOTIENT",
            "RADIANS",
            "RANDBETWEEN",
            "RATE",
            "RECEIVED",
            "REPLACE",
            "REPLACEB",
            "REPT",
            "RIGHT",
            "RIGHTB",
            "ROMAN",
            "ROUND",
            "ROUNDDOWN",
            "ROUNDUP",
            "RRI",
            "SEARCH",
            "SEARCHB",
            "SEC",
            "SECH",
            "SECOND",
            "SIGN",
            "SIN",
            "SINH",
            "SLN",
            "SQRT",
            "SQRTPI",
            "STANDARDIZE",
            "SUBSTITUTE",
            "SYD",
            "T.DIST",
            "T.DIST.2T",
            "T.DIST.RT",
            "T.INV",
            "T.INV.2T",
            "TAN",
            "TANH",
            "TBILLEQ",
            "TBILLPRICE",
            "TBILLYIELD",
            "TDIST",
            "TEXT",
            "TIME",
            "TIMEVALUE",
            "TINV",
            "TRIM",
            "TRUNC",
            "UNICHAR",
            "UNICODE",
            "UPPER",
            "VALUE",
            "WEEKDAY",
            "WEEKNUM",
            "WEIBULL",
            "WEIBULL.DIST",
            "YEAR",
            "YEARFRAC",
            "YIELD",
            "YIELDDISC",
            "YIELDMAT",
        ] {
            table.insert(name, Params::All(Value));
        }
        for name in [
            "AND",
            "AVEDEV",
            "AVERAGE",
            "AVERAGEA",
            "COUNT",
            "COUNTA",
            "COUNTBLANK",
            "DEVSQ",
            "GCD",
            "GEOMEAN",
            "HARMEAN",
            "IMPRODUCT",
            "IMSUM",
            "KURT",
            "LCM",
            "MAX",
            "MAXA",
            "MEDIAN",
            "MIN",
            "MINA",
            "MULTINOMIAL",
            "OR",
            "PRODUCT",
            "SKEW",
            "SKEW.P",
            "STDEV",
            "STDEV.P",
            "STDEV.S",
            "STDEVA",
            "STDEVP",
            "STDEVPA",
            "SUM",
            "SUMSQ",
            "VAR",
            "VAR.P",
            "VAR.S",
            "VARA",
            "VARP",
            "VARPA",
            "XOR",
            "AREAS",
            "COLUMN",
            "COLUMNS",
            "ISREF",
            "N",
            "ROW",
            "ROWS",
            "T",
            "ISFORMULA",
            "FORMULATEXT",
            "SHEET",
            "SHEETS",
            "TYPE",
        ] {
            table.insert(name, Params::All(Reference));
        }
        for name in [
            "CHISQ.TEST",
            "CHITEST",
            "CORREL",
            "COVAR",
            "COVARIANCE.P",
            "COVARIANCE.S",
            "F.TEST",
            "FREQUENCY",
            "FTEST",
            "GROWTH",
            "INTERCEPT",
            "LINEST",
            "LOGEST",
            "MDETERM",
            "MINVERSE",
            "MMULT",
            "MODE",
            "MODE.MULT",
            "MODE.SNGL",
            "PEARSON",
            "PROB",
            "RSQ",
            "SLOPE",
            "STEYX",
            "SUMPRODUCT",
            "SUMX2MY2",
            "SUMX2PY2",
            "SUMXMY2",
            "T.TEST",
            "TRANSPOSE",
            "TREND",
            "TTEST",
        ] {
            table.insert(name, Params::All(Array));
        }
        for name in ["IF", "IFERROR", "IFNA", "CHOOSE"] {
            table.insert(name, Params::Listed(&[Scalar], Branch));
        }
        table.insert("INDEX", Params::Listed(&[Array], Scalar));
        table.insert("LOOKUP", Params::Listed(&[Value], Array));
        table.insert("MATCH", Params::Listed(&[Value, Reference], Value));
        for name in ["VLOOKUP", "HLOOKUP"] {
            table.insert(name, Params::Listed(&[Scalar, Reference], Scalar));
        }
        table.insert("OFFSET", Params::Listed(&[Reference], Value));
        for name in ["COUNTIF", "SUMIF", "AVERAGEIF"] {
            table.insert(name, Params::Listed(&[Reference, Value], Reference));
        }
        table.insert("COUNTIFS", Params::RangeCriteria);
        for name in ["SUMIFS", "AVERAGEIFS", "MAXIFS", "MINIFS"] {
            table.insert(name, Params::SumRangeCriteria);
        }
        for name in [
            "SMALL",
            "LARGE",
            "PERCENTILE",
            "PERCENTILE.EXC",
            "PERCENTILE.INC",
            "QUARTILE",
            "QUARTILE.EXC",
            "QUARTILE.INC",
            "PERCENTRANK",
            "PERCENTRANK.EXC",
            "PERCENTRANK.INC",
            "TRIMMEAN",
        ] {
            table.insert(name, Params::Listed(&[Reference], Value));
        }
        for name in ["RANK", "RANK.EQ", "RANK.AVG"] {
            table.insert(name, Params::Listed(&[Value, Reference], Value));
        }
        table.insert("NPV", Params::Listed(&[Value], Reference));
        for name in ["IRR", "MIRR"] {
            table.insert(name, Params::Listed(&[Array], Value));
        }
        table.insert("XIRR", Params::Listed(&[Reference, Reference], Value));
        table.insert("XNPV", Params::Listed(&[Value, Reference], Reference));
        table.insert("SUBTOTAL", Params::Listed(&[Value], Reference));
        table.insert("AGGREGATE", Params::Listed(&[Value, Value], Reference));
        for name in ["FORECAST", "FORECAST.LINEAR"] {
            table.insert(name, Params::Listed(&[Value], Array));
        }
        table.insert("CELL", Params::Listed(&[Value], Reference));
        table
    })
}

fn class_in(params: Params, at: usize) -> ParamClass {
    match params {
        Params::All(class) => class,
        Params::Listed(first, rest) => first.get(at).copied().unwrap_or(rest),
        Params::RangeCriteria => {
            if at % 2 == 0 {
                ParamClass::Reference
            } else {
                ParamClass::Value
            }
        }
        Params::SumRangeCriteria => {
            if at == 0 || at % 2 == 1 {
                ParamClass::Reference
            } else {
                ParamClass::Value
            }
        }
    }
}

/// A function's name without the prefixes a file gives newer functions.
fn bare(name: &str) -> &str {
    let name = name
        .strip_prefix("_xlfn.")
        .or_else(|| name.strip_prefix("_XLFN."))
        .unwrap_or(name);
    name.strip_prefix("_xlws.")
        .or_else(|| name.strip_prefix("_XLWS."))
        .unwrap_or(name)
}

/// The class of a function's parameter. A function that isn't known here, which takes in everything
/// newer than dynamic arrays, gets `None`: it deals with arrays itself.
pub fn param_class(name: &str, at: usize) -> Option<ParamClass> {
    let name = bare(name);
    let params = match table().get(name) {
        Some(params) => *params,
        None => *table().get(name.to_ascii_uppercase().as_str())?,
    };
    Some(class_in(params, at))
}

/// Whether a function reads an array of references given to this reference parameter one by one, as
/// `SUMIF(INDIRECT({"A!A:A","B!A:A"}),1)` does for a sum over sheets. Most functions can't.
pub(crate) fn reads_references_one_by_one(name: &str, at: usize) -> bool {
    matches!(param_class(name, at), Some(ParamClass::Reference))
        && matches!(
            name,
            "SUMIF"
                | "COUNTIF"
                | "AVERAGEIF"
                | "SUMIFS"
                | "COUNTIFS"
                | "AVERAGEIFS"
                | "MAXIFS"
                | "MINIFS"
                | "SUBTOTAL"
        )
}

/// Whether a function runs once for each value of an array given to a single-value parameter.
/// `IF` and those like it deal with arrays themselves, and `INDEX` has its own rules.
pub(crate) fn lifts_at(name: &str, at: usize) -> bool {
    matches!(
        param_class(name, at),
        Some(ParamClass::Value | ParamClass::Scalar)
    ) && !matches!(
        name,
        "IF" | "IFERROR" | "IFNA" | "CHOOSE" | "INDEX" | "INDIRECT" | "HYPERLINK"
    )
}
