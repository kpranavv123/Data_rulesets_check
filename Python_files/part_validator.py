import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter



# ─────────────────────────────────────────────
#  >>>  ADD YOUR VALID PLANT VALUES HERE  <<<
# ─────────────────────────────────────────────
VALID_PLANTS = [
    "1127", "1100", "1105", "1146", "1156", "1107", "1157", "1158", "1166", "1180",
    "1184", "1186", "1197", "1203", "1204", "1211", "1213", "1214", "1218", "1223",
    "1110", "1225", "1226", "1229", "1233", "1234", "1240", "1113", "1248", "1253",
    "1257", "1258", "1265", "1114", "1145", "1275", "1279", "1416", "1421", "1423",
    "1425", "1426", "1428", "1429", "1430", "1432", "1433", "1436", "1437", "1438",
    "1439", "1440", "1442", "1445", "1449", "1451", "1452", "1455", "1463", "1471",
    "1473", "1475", "1476", "1477", "1478", "1480", "1481", "1483", "1484", "1485",
    "1487", "1488", "1491", "1495", "1500", "1501", "1505", "1506", "1507", "1508",
    "1509", "1521", "1525", "1563", "1578", "1650", "1651", "1652", "1654", "1656",
    "1657", "1658", "1659", "1661", "1579", "1627", "1589", "1623", "1646", "1647",
    "1640", "1112", "5011637", "5123296", "5123742", "4007430", "1722", "1724", "1725",
    "1726", "1731", "1732", "1733", "1734", "1738", "1739", "1740", "1742", "1754",
    "1757", "1758", "1771", "1774", "1780", "1784", "1785", "1788", "1511", "1512",
    "1948", "1448", "4011702", "5013796", "5018849", "5011407", "5015073", "5011308",
    "5123531", "1642", "1638", "1104", "1104A", "1643", "1106", "1109", "1645", "1111",
    "1648", "1649", "1653", "1801", "1802", "2082", "2091", "2088", "2089", "2083",
    "2084", "2085", "2090", "2081", "2086", "2087", "1554", "1520", "1472", "1571",
    "1298", "1295", "1196", "1292", "1176", "1441", "1522", "1494", "1489", "1518",
    "1249", "1296", "1137", "1208", "1155", "1569", "1235", "1281", "1503", "1482",
    "1135", "1205", "1241", "1499", "1462", "1555", "1559", "1575", "2092", "1558",
    "1601", "1125", "1256", "1568",
]


# ──────────────────────────────────────────────
#  Colours
# ──────────────────────────────────────────────
RED_FILL   = PatternFill("solid", start_color="FF0000", end_color="FF0000")
ROW_FILL   = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")   # light yellow
HDR_FILL   = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")   # light blue header
HDR_FONT   = Font(bold=True, name="Arial")
BODY_FONT  = Font(name="Arial", size=10)


# ══════════════════════════════════════════════
#  Rule Engine
# ══════════════════════════════════════════════
class RuleEngine:
    """Contains all column-level validation rules for the Part Table."""

    def __init__(self, valid_plants: list):
        self.valid_plants = [str(p).strip() for p in valid_plants]

    # helpers
    @staticmethod
    def _is_blank(value) -> bool:
        return pd.isna(value) or str(value).strip() == ""

    def _check_not_blank(self, value) -> bool:
        return not self._is_blank(value)

    # individual column validators  →  return True = PASS, False = FAIL
    def validate_material_number(self, row) -> bool:
        mat   = row.get("MATERIALNUMBER")
        ptype = str(row.get("PRODUCTTYPE", "")).strip().upper()
        if self._is_blank(mat):
            return False
        try:
            num = int(float(str(mat).strip()))
        except (ValueError, TypeError):
            return False
        if ptype == "FERT":
            return 14000000000000 <= num <= 14999999999999
        if ptype == "HAWA":
            return 15000000000000 <= num <= 15999999999999
        return False

    def validate_plant(self, row) -> bool:
        val = row.get("PLANT")
        if self._is_blank(val):
            return False
        return str(val).strip() in self.valid_plants

    def validate_product_description(self, row) -> bool:
        return self._check_not_blank(row.get("PRODUCTDESCRIPTION"))

    def validate_product_type(self, row) -> bool:
        val = str(row.get("PRODUCTTYPE", "")).strip().upper()
        return val in {"FERT", "HAWA"}

    def validate_product_hierarchy(self, row) -> bool:
        return self._check_not_blank(row.get("PRODUCTHIERARCHY"))

    def validate_mrp_type(self, row) -> bool:
        val = str(row.get("MRPTYPE", "")).strip().upper()
        return val in {"ND", "PD"}

    def validate_procurement_type(self, row) -> bool:
        return self._check_not_blank(row.get("PROCUREMENTTYPE"))

    def validate_abc_indicator(self, row) -> bool:
        return self._check_not_blank(row.get("ABCINDICATOR"))

    def validate_ibp_status(self, row) -> bool:
     raw = row.get("IBPSTATUS")
     if self._is_blank(raw):
        return True          # blank is allowed → no error
     return str(raw).strip().upper() == "IBP"

    def validate_xplant_mat_status(self, row) -> bool:
     raw = row.get("XPLANTMATSTATUS")
     if self._is_blank(raw):
        return True          # blank is allowed → no error
     return str(raw).strip() == "2"

    # map: column name → validator method
    def get_rules(self) -> dict:
        return {
            "MATERIALNUMBER":    self.validate_material_number,
            "PLANT":             self.validate_plant,
            "PRODUCTDESCRIPTION":self.validate_product_description,
            "PRODUCTTYPE":       self.validate_product_type,
            "PRODUCTHIERARCHY":  self.validate_product_hierarchy,
            "MRPTYPE":           self.validate_mrp_type,
            "PROCUREMENTTYPE":   self.validate_procurement_type,
            "ABCINDICATOR":      self.validate_abc_indicator,
            "IBPSTATUS":         self.validate_ibp_status,
            "XPLANTMATSTATUS":   self.validate_xplant_mat_status,
        }


# ══════════════════════════════════════════════
#  Validator  –  runs rules on the DataFrame
# ══════════════════════════════════════════════
class PartTableValidator:
    """Reads the Excel file, applies rules, and produces an error map."""

    def __init__(self, filepath: str, valid_plants: list):
        self.filepath     = filepath
        self.engine       = RuleEngine(valid_plants)
        self.df           = pd.DataFrame()
        self.error_map    = {}        # {row_index: [col_name, ...]}

    def load(self):
        self.df = pd.read_excel(self.filepath, dtype=str)
        self.df.columns = [c.strip().upper() for c in self.df.columns]

    def validate(self):
        rules = self.engine.get_rules()
        for idx, row in self.df.iterrows():
            failed_cols = []
            for col, rule_fn in rules.items():
                if col not in self.df.columns:
                    continue
                try:
                    passed = rule_fn(row)
                except Exception:
                    passed = False
                if not passed:
                    failed_cols.append(col)
            if failed_cols:
                self.error_map[idx] = failed_cols

    def get_error_summary(self) -> pd.Series:
        """Returns a Series aligned to df index with comma-joined error columns."""
        return pd.Series({
            idx: ", ".join(cols)
            for idx, cols in self.error_map.items()
        }, dtype=str)


# ══════════════════════════════════════════════
#  Report Writer  –  builds the output workbook
# ══════════════════════════════════════════════
class ReportWriter:
    """Writes the validated data to an output Excel workbook with 2 sheets."""

    SHEET_ALL    = "Full Data"
    SHEET_ERRORS = "Error Rows"

    def __init__(self, validator: PartTableValidator, output_path: str):
        self.validator   = validator
        self.output_path = output_path

    # ── internal helpers ──────────────────────
    def _write_dataframe_to_sheet(self, ws, df: pd.DataFrame):
        """Writes DataFrame (with header) into a worksheet."""
        for c_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=1, column=c_idx, value=col_name)
            cell.fill = HDR_FILL
            cell.font = HDR_FONT

        for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
            for c_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.font = BODY_FONT

    def _apply_highlights(self, ws, df: pd.DataFrame, error_map: dict, col_index: dict):
        """
        Highlights error rows with light yellow and individual error cells with red.
        df_row_idx (0-based) → excel_row = df_row_idx + 2  (accounting for header)
        """
        for df_idx, bad_cols in error_map.items():
            excel_row = df_idx + 2          # +1 header, +1 because openpyxl is 1-based
            total_cols = len(df.columns)

            # light-yellow entire row
            for c in range(1, total_cols + 1):
                ws.cell(row=excel_row, column=c).fill = ROW_FILL

            # red on specific error cells
            for col_name in bad_cols:
                if col_name in col_index:
                    ws.cell(row=excel_row, column=col_index[col_name]).fill = RED_FILL
                    ws.cell(row=excel_row, column=col_index[col_name]).font = Font(
                        name="Arial", size=10, bold=True, color="FFFFFF"
                    )

    def _set_column_widths(self, ws):
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value else 0) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 50)

    # ── public ───────────────────────────────
    def write(self):
        v      = self.validator
        df     = v.df.copy()
        errors = v.error_map

        # add ERROR_COLUMNS summary column
        error_series = v.get_error_summary()
        df["ERROR_COLUMNS"] = df.index.map(lambda i: error_series.get(i, ""))

        # build column → excel-column-index map (1-based)
        col_index = {col: i for i, col in enumerate(df.columns, start=1)}

        # error-only rows (preserve original row numbers via index)
        error_df = df[df.index.isin(errors.keys())].copy()

        # ── write both sheets ─────────────────
        # We write via openpyxl directly for formatting control
        from openpyxl import Workbook
        wb = Workbook()

        # Sheet 1: Full Data
        ws_all = wb.active
        ws_all.title = self.SHEET_ALL
        self._write_dataframe_to_sheet(ws_all, df)
        self._apply_highlights(ws_all, df, errors, col_index)
        self._set_column_widths(ws_all)

        # Sheet 2: Error Rows only
        ws_err = wb.create_sheet(self.SHEET_ERRORS)
        self._write_dataframe_to_sheet(ws_err, error_df)

        # re-map errors to error_df sequential positions for sheet 2
        err_col_idx = {col: i for i, col in enumerate(error_df.columns, start=1)}
        for sheet2_row, orig_idx in enumerate(error_df.index, start=2):
            bad_cols = errors[orig_idx]
            total_cols = len(error_df.columns)
            for c in range(1, total_cols + 1):
                ws_err.cell(row=sheet2_row, column=c).fill = ROW_FILL
            for col_name in bad_cols:
                if col_name in err_col_idx:
                    ws_err.cell(row=sheet2_row, column=err_col_idx[col_name]).fill = RED_FILL
                    ws_err.cell(row=sheet2_row, column=err_col_idx[col_name]).font = Font(
                        name="Arial", size=10, bold=True, color="FFFFFF"
                    )

        self._set_column_widths(ws_err)
        wb.save(self.output_path)
        print(f"\n✅  Output saved → {self.output_path}")
        print(f"   Total rows     : {len(df)}")
        print(f"   Error rows     : {len(error_df)}")


# ══════════════════════════════════════════════
#  Orchestrator
# ══════════════════════════════════════════════
class PartTableProcessor:
    """Ties together loading, validation, and report writing."""

    def __init__(self, input_path: str, output_path: str, valid_plants: list):
        self.validator = PartTableValidator(input_path, valid_plants)
        self.writer    = ReportWriter(self.validator, output_path)

    def run(self):
        print("📂  Loading file …")
        self.validator.load()
        print(f"    Columns detected: {list(self.validator.df.columns)}")

        print("🔍  Validating rules …")
        self.validator.validate()

        print("📝  Writing report …")
        self.writer.write()


# ══════════════════════════════════════════════
#  Entry Point
# ══════════════════════════════════════════════
if __name__ == "__main__":

    INPUT_FILE  = r"D:/SEM-8/Data Rules Set Check/Data_rulesets_check/Excel_Files/Part.xlsx"   # ← change this
    OUTPUT_FILE = r"D:/SEM-8/Data Rules Set Check/Data_rulesets_check/Output_Files/Validated_Part.xlsx"  # ← change this

    processor = PartTableProcessor(
        input_path   = INPUT_FILE,
        output_path  = OUTPUT_FILE,
        valid_plants = VALID_PLANTS,
    )
    processor.run()