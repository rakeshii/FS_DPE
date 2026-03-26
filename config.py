"""
config.py — All configuration in one place.
Change values here; nothing else needs editing.
"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


class Config:
    # ── Security ──────────────────────────────────────────────
    SECRET_KEY = os.environ.get('SECRET_KEY', 'dev-key-change-in-production')

    # ── File handling ─────────────────────────────────────────
    MAX_CONTENT_LENGTH = 20 * 1024 * 1024          # 20 MB upload limit
    ALLOWED_EXTENSIONS = {'xls', 'xlsx'}
    # Use /tmp on cloud (Railway) — always writable. Falls back to project dirs locally.
    UPLOAD_FOLDER      = os.environ.get('UPLOAD_DIR', '/tmp/fsproj_uploads')
    OUTPUT_FOLDER      = os.environ.get('OUTPUT_DIR', '/tmp/fsproj_outputs')

    # ── LibreOffice ───────────────────────────────────────────
    # Override with env var if libreoffice is in a custom path
    LIBREOFFICE_CMD = os.environ.get('LIBREOFFICE_CMD', 'libreoffice')
    LIBREOFFICE_TIMEOUT = 90   # seconds

    # ── Projector defaults ────────────────────────────────────
    DEFAULT_NEW_HEADER  = 'As at 31 March, 2026'
    DEFAULT_OUTPUT_NAME = 'FS-FY_2025-26_Draft.xlsx'

    # ── Financial sheet config (1-indexed insert column) ─────
    #    insert = column number where the new 2026 col is inserted
    SHEET_CONFIG = {
        'BS':               {'insert': 4},
        'P & L':            {'insert': 4},
        'CFS':              {'insert': 4},
        'Note 2':           {'insert': 4, 'width': 2},
        'Note  3-4':        {'insert': 4},
        'Note  5-7':        {'insert': 4},
        'Note 8':           {'insert': 6},
        'Note 9-11':        {'insert': 4},
        'Note 11-14':       {'insert': 4},
        'Note To P & L':    {'insert': 3},
        'Details to notes': {'insert': 2},
    }

    # Merged header chain patches
    # Format: (sheet_name, row_1idx, col_1idx, original_formula)
    CHAIN_PATCHES = [
        ('Note 2',     2, 2, "=+'Note  3-4'!B2:E2"),
        ('Note 2',     3, 2, "=+'Note  3-4'!B3:E3"),
        ('Note  3-4',  2, 2, "='P & L'!B2:E2"),
        ('Note  5-7',  2, 2, "='P & L'!B2:E2"),
        ('Note 8',     4, 2, "=+'Note  5-7'!B37:E37"),
        ('Note 9-11',  2, 2, "=+'Note 8'!B4:E4"),
        ('Note 11-14', 2, 2, "=+'Note 9-11'!B2:E2"),
        ('Note 11-14', 3, 2, "=+'Note 9-11'!B3:E3"),
    ]



    # ── TB Named Ranges ───────────────────────────────────────────────────
    # Maps every TB row referenced by financial sheets to a stable name.
    # Format: { 'Name': (row_number, 'B'|'C'|'BC') }
    # 'BC' = create both TB_B_ and TB_C_ versions
    # 'B'  = Debit column only
    # 'C'  = Credit column only
    TB_NAMED_RANGES = {
        'YES_Bank_OD':             (21,  'BC'),
        'Yes_Bank_Margin':         (22,  'BC'),
        'Unsecured_Ashish':        (24,  'BC'),
        'Unsecured_Ashish_HUF':    (25,  'BC'),
        'Unsecured_Dev':           (26,  'BC'),
        'Unsecured_Devakinandan':  (27,  'BC'),
        'Unsecured_Rashi':         (28,  'BC'),
        'Unsecured_Seema':         (29,  'BC'),
        'HDFC_Credit_Card':        (30,  'BC'),
        'CGST_Payable':            (34,  'BC'),
        'TCS_Collected':           (45,  'C'),
        'TDS_194J':                (48,  'C'),
        'Audit_Fee_Payable':       (56,  'C'),
        'Electricity_Payable':     (57,  'C'),
        'Salaries_Payable':        (58,  'C'),
        'Telephone_Payable':       (59,  'C'),
        'Sundry_Creditors':        (60,  'BC'),
        'Electricity_Deposit':     (80,  'B'),
        'Sundry_Debtors':          (81,  'BC'),
        'Cash':                    (83,  'B'),
        'Advances_Expenses':       (85,  'B'),
        'IT_Refund':               (86,  'B'),
        'IT_Refund_2324':          (87,  'B'),
        'Prepaid_Insurance':       (88,  'B'),
        'TDS_FD_Interest':         (90,  'B'),
        'TDS_Making_2pct':         (91,  'B'),
        'TDS_Hallmarking':         (92,  'B'),
        'TDS_Sale_01pct':          (93,  'B'),
        'Advance_Tax':             (94,  'B'),
        'Deferred_Tax':            (95,  'B'),
        'Sales_Jobwork_Out':       (97,  'C'),
        'Sales_Jobwork':           (98,  'C'),
        'Sales_Rate_Diff':         (99,  'C'),
        'Sales':                   (100, 'C'),
        'Purchase_Customs':        (102, 'C'),
        'Purchase_Rate_Diff':      (103, 'B'),
        'Purchase':                (104, 'B'),
        'Purchase_Rate_Diff2':     (105, 'BC'),
        'Purchase_Trans_Diff':     (106, 'B'),
        'Making_Charges':          (109, 'B'),
        'Making_Charges_JW':       (110, 'B'),
        'Discount':                (112, 'C'),
        'Forex_Gain_Loss':         (113, 'C'),
        'Hallmarking_IGST':        (114, 'C'),
        'Hallmarking_Inc':         (115, 'C'),
        'Interest_FD':             (116, 'C'),
        'CSR_Raginiben':           (119, 'B'),
        'Director_Salary_Ashish':  (123, 'B'),
        'Director_Salary_Seema':   (124, 'B'),
        'Staff_B_Gnaneshwar':      (126, 'B'),
        'Bank_Charges':            (171, 'B'),
        'Interest_Metal_Loan':     (172, 'B'),
        'Interest_Unsecured':      (173, 'B'),
        'Interest_Yes_Bank':       (174, 'B'),
        'Advertisement_Exp':       (175, 'B'),
        'Agency_Charges':          (176, 'B'),
        'Audit_Fee':               (177, 'B'),
        'Brokerage':               (178, 'B'),
        'Business_Promotion':      (179, 'B'),
        'Depreciation':            (180, 'B'),
        'Diesel_Petrol':           (181, 'B'),
        'Electricity_Exp':         (182, 'B'),
        'Exhibition_Exp':          (183, 'B'),
        'Extinguishment':          (184, 'B'),
        'Freight':                 (185, 'B'),
        'GST_Exp':                 (186, 'B'),
        'Hallmarking_Exp':         (187, 'B'),
        'Insurance_Exp':           (188, 'B'),
        'Interest_TCS_Late':       (189, 'B'),
        'Internal_Audit':          (190, 'B'),
        'Office_Exp':              (191, 'B'),
        'Other_Charges':           (192, 'B'),
        'Parcel_Charges':          (193, 'B'),
        'Printing_Stationary':     (194, 'B'),
        'Professional_Fees':       (195, 'B'),
        'Registrations':           (196, 'B'),
        'Remittance_Charges':      (197, 'B'),
        'Rent':                    (198, 'B'),
        'Repairs_Maintenance':     (199, 'B'),
        'Round_Off':               (200, 'C'),
        'Security_Service':        (201, 'B'),
        'Software_Exp':            (202, 'B'),
        'Staff_Welfare':           (203, 'B'),
        'Subscription':            (204, 'B'),
        'Sundry_Written_Off':      (205, 'B'),
        'Telephone_Exp':           (206, 'B'),
        'Travelling_Exp':          (207, 'B'),
        'Vaulting':                (208, 'B'),
        'Vehicle_Insurance':       (209, 'B'),
        'Vehicle_Maintenance':     (210, 'B'),
        'Water_Sewage':            (211, 'B'),
        # Salary block range (rows 126-169) — used as a range sum
        'All_Salaries_B':          ((126, 169), 'B'),
        'All_Salaries_C':          ((126, 169), 'C'),
        # TDS block ranges (Note 5-7 R49,50)
        'TDS_Block_45_47_C':       ((45, 47),   'C'),
        'TDS_Block_45_47_B':       ((45, 47),   'B'),
        'TDS_Block_45_46_B':       ((45, 46),   'B'),
        'TDS_Block_48_54_C':       ((48, 54),   'C'),
        'TDS_Block_48_54_B':       ((48, 54),   'B'),
        # Purchase block (Note To P&L R32)
        'Purchase_Block_102_106_C':((102, 106), 'C'),
        'Purchase_Block_102_106_B':((102, 106), 'B'),
        # CSR block (Note To P&L)
        'CSR_Block_119_120_B':     ((119, 120), 'B'),
        # Making charges block
        'Making_Block_109_110_B':  ((109, 110), 'B'),
        # Sales block
        'Sales_Block_97_100_C':    ((97, 100),  'C'),
    }


class DevelopmentConfig(Config):
    DEBUG = True
    TESTING = False


class ProductionConfig(Config):
    DEBUG = False
    TESTING = False
    SECRET_KEY = os.environ.get('SECRET_KEY')   # Must be set in env


class TestingConfig(Config):
    TESTING = True
    DEBUG = True
