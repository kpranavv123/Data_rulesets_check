import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
#  FILE PATHS  –  update these
# ─────────────────────────────────────────────
SITE_INPUT_FILE  = r"D:\SEM-8\Data Rules Set Check\Data_rulesets_check\Excel_Files\Site.xlsx"
PART_INPUT_FILE  = r"D:\SEM-8\Data Rules Set Check\Data_rulesets_check\Excel_Files\Part.xlsx"
OUTPUT_FILE      = r"D:\SEM-8\Data Rules Set Check\Data_rulesets_check\Output_filesValidated_Site.xlsx"


# ─────────────────────────────────────────────
#  CONSOLIDATED PL LIST
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


# ─────────────────────────────────────────────
#  Colours
# ─────────────────────────────────────────────
RED_FILL  = PatternFill("solid", start_color="FF0000", end_color="FF0000")
ROW_FILL  = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
HDR_FILL  = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
HDR_FONT  = Font(bold=True, name="Arial")
BODY_FONT = Font(name="Arial", size=10)
ERR_FONT  = Font(name="Arial", size=10, bold=True, color="FFFFFF")


# ══════════════════════════════════════════════
#  Rule Engine
# ══════════════════════════════════════════════
class SiteRuleEngine:
    """All column-level validation rules for the Site Table."""

    def __init__(self, valid_plants: list, part_plants: set):
        self.valid_plants = set(str(p).strip() for p in valid_plants)
        self.part_plants  = set(str(p).strip() for p in part_plants)

    @staticmethod
    def _is_blank(value) -> bool:
        return pd.isna(value) or str(value).strip() == ""

    def validate_plant(self, row) -> bool:
        val = str(row.get("PLANT", "")).strip()
        if not val:
            return False
        in_pl_list  = val in self.valid_plants
        in_part_tbl = val in self.part_plants
        return in_pl_list and in_part_tbl

    def validate_name(self, row) -> bool:
        return not self._is_blank(row.get("NAME"))

    def validate_address(self, row) -> bool:
        return not self._is_blank(row.get("ADDRESS"))

    def validate_tcpl_planttype(self, row) -> bool:
        return not self._is_blank(row.get("TCPL_PLANTTYPE"))

    def get_rules(self) -> dict:
        return {
            "PLANT":          self.validate_plant,
            "NAME":           self.validate_name,
            "ADDRESS":        self.validate_address,
            "TCPL_PLANTTYPE": self.validate_tcpl_planttype,
        }


# ══════════════════════════════════════════════
#  Validator
# ══════════════════════════════════════════════
class SiteTableValidator:
    """Loads Site & Part Excel files, runs validation, builds error map."""

    def __init__(self, site_path: str, part_path: str, valid_plants: list):
        self.site_path   = site_path
        self.part_path   = part_path
        self.valid_plants = valid_plants
        self.df          = pd.DataFrame()
        self.error_map   = {}

    def load(self):
        self.df = pd.read_excel(self.site_path, dtype=str)
        self.df.columns = [c.strip().upper() for c in self.df.columns]

        part_df = pd.read_excel(self.part_path, dtype=str)
        part_df.columns = [c.strip().upper() for c in part_df.columns]

        if "PLANT" not in part_df.columns:
            raise ValueError("PLANT column not found in Part table.")

        self.part_plants = set(part_df["PLANT"].dropna().str.strip().tolist())
        print(f"    Part table plants loaded  : {len(self.part_plants)} unique values")

    def validate(self):
        engine = SiteRuleEngine(self.valid_plants, self.part_plants)
        rules  = engine.get_rules()

        for idx, row in self.df.iterrows():
            failed = []
            for col, rule_fn in rules.items():
                if col not in self.df.columns:
                    continue
                try:
                    passed = rule_fn(row)
                except Exception:
                    passed = False
                if not passed:
                    failed.append(col)
            if failed:
                self.error_map[idx] = failed

    def get_error_series(self) -> pd.Series:
        return pd.Series(
            {idx: ", ".join(cols) for idx, cols in self.error_map.items()},
            dtype=str
        )


# ══════════════════════════════════════════════
#  Report Writer
# ══════════════════════════════════════════════
class SiteReportWriter:
    """Builds the 2-sheet output Excel with colour highlights."""

    SHEET_ALL    = "Full Data"
    SHEET_ERRORS = "Error Rows"

    def __init__(self, validator: SiteTableValidator, output_path: str):
        self.validator   = validator
        self.output_path = output_path

    def _write_header(self, ws, columns):
        for c_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=1, column=c_idx, value=col_name)
            cell.fill = HDR_FILL
            cell.font = HDR_FONT

    def _write_rows(self, ws, df: pd.DataFrame):
        for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
            for c_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.font = BODY_FONT

    def _highlight(self, ws, df: pd.DataFrame, error_map: dict, col_index: dict):
        for df_idx, bad_cols in error_map.items():
            excel_row  = df_idx + 2
            total_cols = len(df.columns)
            for c in range(1, total_cols + 1):
                ws.cell(row=excel_row, column=c).fill = ROW_FILL
            for col_name in bad_cols:
                if col_name in col_index:
                    cell = ws.cell(row=excel_row, column=col_index[col_name])
                    cell.fill = RED_FILL
                    cell.font = ERR_FONT

    def _set_widths(self, ws):
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value else 0) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 50)

    def write(self):
        v  = self.validator
        df = v.df.copy()
        df["ERROR_COLUMNS"] = df.index.map(lambda i: v.get_error_series().get(i, ""))

        col_index = {col: i for i, col in enumerate(df.columns, start=1)}
        error_df  = df[df.index.isin(v.error_map.keys())].copy()

        wb = Workbook()

        # ── Sheet 1: Full Data ────────────────
        ws_all = wb.active
        ws_all.title = self.SHEET_ALL
        self._write_header(ws_all, df.columns)
        self._write_rows(ws_all, df)
        self._highlight(ws_all, df, v.error_map, col_index)
        self._set_widths(ws_all)

        # ── Sheet 2: Error Rows ───────────────
        ws_err = wb.create_sheet(self.SHEET_ERRORS)
        self._write_header(ws_err, error_df.columns)
        self._write_rows(ws_err, error_df)

        err_col_idx = {col: i for i, col in enumerate(error_df.columns, start=1)}
        for sheet2_row, orig_idx in enumerate(error_df.index, start=2):
            bad_cols   = v.error_map[orig_idx]
            total_cols = len(error_df.columns)
            for c in range(1, total_cols + 1):
                ws_err.cell(row=sheet2_row, column=c).fill = ROW_FILL
            for col_name in bad_cols:
                if col_name in err_col_idx:
                    cell = ws_err.cell(row=sheet2_row, column=err_col_idx[col_name])
                    cell.fill = RED_FILL
                    cell.font = ERR_FONT

        self._set_widths(ws_err)
        wb.save(self.output_path)

        print(f"\n✅  Output saved  → {self.output_path}")
        print(f"   Total rows    : {len(df)}")
        print(f"   Error rows    : {len(error_df)}")


# ══════════════════════════════════════════════
#  Orchestrator
# ══════════════════════════════════════════════
class SiteTableProcessor:
    """Ties together loading, validation, and report writing."""

    def __init__(self, site_path: str, part_path: str, output_path: str, valid_plants: list):
        self.validator = SiteTableValidator(site_path, part_path, valid_plants)
        self.writer    = SiteReportWriter(self.validator, output_path)

    def run(self):
        print("📂  Loading files …")
        self.validator.load()
        print(f"    Site columns detected : {list(self.validator.df.columns)}")

        print("🔍  Validating rules …")
        self.validator.validate()

        print("📝  Writing report …")
        self.writer.write()


# ══════════════════════════════════════════════
#  Entry Point
# ══════════════════════════════════════════════
if __name__ == "__main__":
    processor = SiteTableProcessor(
        site_path   = SITE_INPUT_FILE,
        part_path   = PART_INPUT_FILE,
        output_path = OUTPUT_FILE,
        valid_plants= VALID_PLANTS,
    )
    processor.run()