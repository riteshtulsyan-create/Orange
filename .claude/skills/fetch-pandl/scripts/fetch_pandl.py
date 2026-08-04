"""
Fetch Profit & Loss from Tally (port 9010) for all open companies.
Writes one Excel sheet per company, matching PandL.xlsx structure.

Usage:
    python fetch_pandl.py --from-date 01-04-2025 --to-date 31-03-2026 --output Files/PandL_Output.xlsx
"""

import argparse
import calendar
import os
import sys
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from datetime import datetime, date
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Load .env: check CWD first (project root), then script dir, then parent
_env_candidates = [
    Path.cwd() / ".env",
    Path(__file__).parent / ".env",
    Path(__file__).parent.parent / ".env",
]
for _env_path in _env_candidates:
    if _env_path.exists():
        for _line in _env_path.read_text().splitlines():
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())
        break

TALLY_HOST = os.environ.get("TALLY_HOST", "localhost")
TALLY_PORT = int(os.environ.get("TALLY_PORT", "9000"))

ACCOUNT_ORDER = [
    (12, "Trading Account:", "group_c", "C"),
    (13, "Sales Accounts", "group_c", "C"),
    (14, "DISCOUNT AND RATE DIFFERENCE (GST)", "ledger", "B"),
    (15, "CONSUMPTION CHARGES (INCOME) 18%", "ledger", "B"),
    (16, "DISCOUNT & RATE DIFFERENCE@12% (PAPER)", "ledger", "B"),
    (17, "DISCOUNT & RATE DIFFERENCE@18%-HEAD", "ledger", "B"),
    (18, "DISCOUNT & RATE DIFFERENCE@18% (INK)", "ledger", "B"),
    (19, "DISCOUNT & RATE DIFFERENCE@18% (MACHINE)", "ledger", "B"),
    (20, "DISCOUNT & RATE DIFFERENCE@18% (PAPER)", "ledger", "B"),
    (21, "DISCOUNT & RATE DIFFERENCE@18%-SPARE", "ledger", "B"),
    (22, "RATE DIFFERENCE (INCOME) 18%-HEAD", "ledger", "B"),
    (23, "RATE DIFFERENCE (INCOME) 18% (INK)", "ledger", "B"),
    (24, "RATE DIFFERENCE (INCOME) 18%-MACHINE", "ledger", "B"),
    (25, "RATE DIFFERENCE (INCOME) 18%-PAPER", "ledger", "B"),
    (26, "RATE DIFFERENCE (INCOME) 18%-SPARE", "ledger", "B"),
    (27, "DISCOUNT AND RATE DIFFERENCE (NON GST)", "ledger", "B"),
    (28, "DISCOUNT", "ledger", "B"),
    (29, "DISCOUNT (BD)", "ledger", "B"),
    (30, "DISCOUNT ON SALE(NON-GST)", "ledger", "B"),
    (31, "DISCOUNT & RATE DIFFERENCE (NON GST)-MACHINE", "ledger", "B"),
    (32, "RELATED PARTY SALES", "ledger", "B"),
    (33, "RELATED PARTY SALES-HEAD", "ledger", "B"),
    (34, "RELATED PARTY SALES-INK", "ledger", "B"),
    (35, "RELATED PARTY SALES-MACHINE", "ledger", "B"),
    (36, "RELATED PARTY SALES-PACKING MATERIAL", "ledger", "B"),
    (37, "RELATED PARTY SALES-PAPER", "ledger", "B"),
    (38, "RELATED PARTY SALES RETURN-HEAD", "ledger", "B"),
    (39, "RELATED PARTY SALES RETURN-INK", "ledger", "B"),
    (40, "RELATED PARTY SALES RETURN-SPARE", "ledger", "B"),
    (41, "RELATED PARTY SALES-SPARE", "ledger", "B"),
    (42, "GST BRANCH SALE ACCOUNTS", "ledger", "B"),
    (43, "GST INTERSTATE SALE - HEAD", "ledger", "B"),
    (44, "GST INTERSTATE SALE -INK", "ledger", "B"),
    (45, "GST INTERSTATE SALE- MACHINE", "ledger", "B"),
    (46, "GST INTERSTATE SALE-PAPER", "ledger", "B"),
    (47, "GST INTERSTATE SALE-SPARE", "ledger", "B"),
    (48, "GST SALES ACCOUNTS -HEAD", "ledger", "B"),
    (49, "GST SALES ACCOUNTS - INK", "ledger", "B"),
    (50, "GST SALES ACCOUNTS - MACHINE", "ledger", "B"),
    (51, "GST SALES ACCOUNTS-PAPER", "ledger", "B"),
    (52, "GST SALES ACCOUNTS - SPARE", "ledger", "B"),
    (53, "GST SALES RETURN-HEAD", "ledger", "B"),
    (54, "GST SALES RETURN-INK", "ledger", "B"),
    (55, "GST SALES RETURN-MACHINE", "ledger", "B"),
    (56, "GST SALES RETURN-PAPER", "ledger", "B"),
    (57, "GST SALES RETURN-SPARE", "ledger", "B"),
    (58, "GST SALES RETURN-SPARE_NON GST", "ledger", "B"),
    (59, "HIGH SEAS SALES", "ledger", "B"),
    (60, "SALES EXPORT", "ledger", "B"),
    (61, "Sales Bills to Make", "group_b", "B"),
    (62, "Direct Incomes", "group_c", "C"),
    (63, "COURIER INCOME", "ledger", "B"),
    (64, "COURIER CHARGES 5%", "ledger", "B"),
    (65, "COURIER CHARGES (INC)", "ledger", "B"),
    (66, "FREIGHT INCOME", "ledger", "B"),
    (67, "CLEARING & FORWERDING CHARGES CENTRING (INCOME)-18%", "ledger", "B"),
    (68, "CLEARING & FORWERDING CHARGES HEAD- (INCOME)-18%", "ledger", "B"),
    (69, "CLEARING & FORWERDING CHARGES INK- (INCOME)-18%", "ledger", "B"),
    (70, "CLEARING & FORWERDING CHARGES MACHINE- (INCOME)-18%", "ledger", "B"),
    (71, "FREIGHT CHARGES@18% (INC)", "ledger", "B"),
    (72, "TRANSPORT CHARGES (INC)", "ledger", "B"),
    (73, "OTHER INCOME", "ledger", "B"),
    (74, "ANNUAL MAINTENANCE CONTRACT (INC)", "ledger", "B"),
    (75, "HANDLING CHARGES ACCOUNTS (INC)", "ledger", "B"),
    (76, "SERVICE CHARGES ACCOUNTS (INC)", "ledger", "B"),
    (77, "", "subtotal_c", "C"),
    (78, "Cost of Sales :", "group_c", "C"),
    (79, "Opening Stock", "ledger", "B"),
    (80, "Add: Purchase Accounts", "ledger", "B"),
    (81, "Less: Closing Stock", "ledger", "B"),
    (82, "", "subtotal_b", "B"),
    (83, "Direct Expenses", "group_b", "B"),
    (84, "CARRIAGE INWARD EXP", "ledger", "B"),
    (85, "CARRIAGE INWARD ACCOUNTS - 12%", "ledger", "B"),
    (86, "CARRIAGE INWARD EXP ACCOUNTS-18%", "ledger", "B"),
    (87, "CARRIAGE INWARD EXP ACCOUNTS@5%", "ledger", "B"),
    (88, "CARRIAGE INWARD EXP ACCOUNTS-NON GST", "ledger", "B"),
    (89, "CARRIAGE INWARD EXP ACCOUNTS (RCM)-FABRIC", "ledger", "B"),
    (90, "CARRIAGE INWARD EXP ACCOUNTS (RCM)-HEAD", "ledger", "B"),
    (91, "CARRIAGE INWARD EXP ACCOUNTS (RCM)-INK", "ledger", "B"),
    (92, "CARRIAGE INWARD EXP ACCOUNTS (RCM)-MACHINE", "ledger", "B"),
    (93, "CARRIAGE INWARD EXP ACCOUNTS (RCM)-PAPER", "ledger", "B"),
    (94, "CARRIAGE INWARD EXP ACCOUNTS (RCM)-POWDER", "ledger", "B"),
    (95, "CARRIAGE INWARD EXP ACCOUNTS (RCM)-SPARE", "ledger", "B"),
    (96, "CLEARING & FORWERDING EXP", "ledger", "B"),
    (97, "CLEARING AND TRANSPORT CHARGES", "ledger", "B"),
    (98, "CLEARING & FORWERDING CHARGES (NON-GST)-HEAD", "ledger", "B"),
    (99, "CLEARING & FORWERDING CHARGES (NON-GST)-INK", "ledger", "B"),
    (100, "CLEARING & FORWERDING CHARGES (NON-GST)-MACHINE", "ledger", "B"),
    (101, "CLEARING & FORWERDING CHARGES (NON-GST)-PAPER", "ledger", "B"),
    (102, "CLEARING & FORWERDING CHARGES (NON-GST)-POWDER", "ledger", "B"),
    (103, "CLEARING & FORWERDING CHARGES (NON-GST)-SPARE", "ledger", "B"),
    (104, "CLEARING & FORWERDING CHARGES-RCM", "ledger", "B"),
    (105, "CLEARNING & FORWERDING CHARGES (12%)-PAPER", "ledger", "B"),
    (106, "CLEARNING & FORWERDING CHARGES (5%)-PAPER", "ledger", "B"),
    (107, "CLEARNING & FORWERDING CHARGES (ITC)-FABRIC", "ledger", "B"),
    (108, "CLEARNING & FORWERDING CHARGES (ITC)-HEAD", "ledger", "B"),
    (109, "CLEARNING & FORWERDING CHARGES (ITC)-INK", "ledger", "B"),
    (110, "CLEARNING & FORWERDING CHARGES (ITC)-MACHINE", "ledger", "B"),
    (111, "CLEARNING & FORWERDING CHARGES (ITC)-MACHINE-EXPORT", "ledger", "B"),
    (112, "CLEARNING & FORWERDING CHARGES (ITC)-PAPER", "ledger", "B"),
    (113, "CLEARNING & FORWERDING CHARGES (ITC)-POWDER", "ledger", "B"),
    (114, "CLEARNING & FORWERDING CHARGES (ITC)-SPARE", "ledger", "B"),
    (115, "CUSTOM DUTY EXPENSES", "ledger", "B"),
    (116, "CUSTOM DUTY EXP ACCOUNTS-FABRICS", "ledger", "B"),
    (117, "CUSTOM DUTY EXP ACCOUNTS-HEAD", "ledger", "B"),
    (118, "CUSTOM DUTY EXP ACCOUNTS-INK", "ledger", "B"),
    (119, "CUSTOM DUTY EXP ACCOUNTS-MACHINE", "ledger", "B"),
    (120, "CUSTOM DUTY EXP ACCOUNTS-PAPER", "ledger", "B"),
    (121, "CUSTOM DUTY EXP ACCOUNTS-SPARE", "ledger", "B"),
    (122, "INTEREST ON DUTY", "ledger", "B"),
    (123, "INTEREST ON DUTY-PAPER", "ledger", "B"),
    (124, "EXCHANGE RATE Gain/(Loss)", "ledger", "B"),
    (125, "EXCHANGE RATE DIFF EXP ACCOUNTS (PURCHASE)", "ledger", "B"),
    (126, "EXCHANGE RATE DIFF EXP ACCOUNTS (PURCHASE)-OLD", "ledger", "B"),
    (127, "EXCHANGE RATE DIFF EXP ACCOUNTS (RECEIPT)", "ledger", "B"),
    (128, "EXCHANGE RATE DIFF EXP ACCOUNTS_LC-BC", "ledger", "B"),
    (129, "EXCHANGE RATE DIFF EXP AC-MAR-26", "ledger", "B"),
    (130, "FREIGHT EXP", "ledger", "B"),
    (131, "AIR FREIGHT", "ledger", "B"),
    (132, "AIR FREIGHT CHARGES @ 5%-SPARE", "ledger", "B"),
    (133, "AIR FREIGHT CHARGES (NON-GST)-HEAD", "ledger", "B"),
    (134, "AIR FREIGHT CHARGES (NON-GST)-INK", "ledger", "B"),
    (135, "AIR FREIGHT CHARGES (NON-GST)-PAPER", "ledger", "B"),
    (136, "AIR FREIGHT CHARGES (NON-GST)-SPARE", "ledger", "B"),
    (137, "AIR FREIGHT EXP ACCOUNTS@18%", "ledger", "B"),
    (138, "AIR FREIGHT EXP ACCOUNTS@18%-HEAD", "ledger", "B"),
    (139, "AIR FREIGHT EXP ACCOUNTS@18% (INK)", "ledger", "B"),
    (140, "AIR FREIGHT EXP ACCOUNTS@18%-PAPER", "ledger", "B"),
    (141, "AIR FREIGHT EXP ACCOUNTS@18%-SPARE", "ledger", "B"),
    (142, "AIR FREIGHT EXP ACCOUNTS(INK)- Export", "ledger", "B"),
    (143, "OCEAN FREIGHT", "ledger", "B"),
    (144, "OCEAN FREIGHT CHARGES @ 18% (HEAD)", "ledger", "B"),
    (145, "OCEAN FREIGHT CHARGES @ 5%-(INK)", "ledger", "B"),
    (146, "OCEAN FREIGHT CHARGES @ 5% (MACHINE)", "ledger", "B"),
    (147, "OCEAN FREIGHT CHARGES @ 5%-(PAPER)", "ledger", "B"),
    (148, "OCEAN FREIGHT CHARGES @ 5% (POWDER)", "ledger", "B"),
    (149, "OCEAN FREIGHT CHARGES @ 5% (SPARE)", "ledger", "B"),
    (150, "OCEAN FREIGHT CHARGES (MACHINE)18%- Export", "ledger", "B"),
    (151, "OCEAN FREIGHT CHARGES (MACHINE)5%- Export", "ledger", "B"),
    (152, "OCEAN FREIGHT CHARGES (MACHINE)- Export", "ledger", "B"),
    (153, "OCEAN FREIGHT EXP ACCOUNTS@12%-MACHINE", "ledger", "B"),
    (154, "OCEAN FREIGHT EXP ACCOUNTS@18% (INK)", "ledger", "B"),
    (155, "OCEAN FREIGHT EXP ACCOUNTS@18% (INK)-EXPORT", "ledger", "B"),
    (156, "OCEAN FREIGHT EXP ACCOUNTS@18% (MACHINE)", "ledger", "B"),
    (157, "OCEAN FREIGHT EXP ACCOUNTS@18%(PAPER)", "ledger", "B"),
    (158, "OCEAN FREIGHT EXP ACCOUNTS@18% (POWDER)", "ledger", "B"),
    (159, "OCEAN FREIGHT EXP ACCOUNTS@18% (SPARE)", "ledger", "B"),
    (160, "OCEAN FREIGHT EXP ACCOUNTS (NON GST)-INK", "ledger", "B"),
    (161, "OCEAN FREIGHT EXP ACCOUNTS (NON GST)-MACHINE", "ledger", "B"),
    (162, "OCEAN FREIGHT EXP ACCOUNTS (NON GST)-(PAPER)", "ledger", "B"),
    (163, "TRANSPORT CHARGES  (PURCHASE)", "ledger", "B"),
    (164, "FREIGHT CHARGE- (12%)", "ledger", "B"),
    (165, "FREIGHT CHARGES (18%)", "ledger", "B"),
    (166, "FREIGHT CHARGES-NON GST", "ledger", "B"),
    (167, "JOB WORK CHARGES", "ledger", "B"),
    (168, "JOB WORK EXP@18% (Related)", "ledger", "B"),
    (169, "JOB WORK EXP@5%", "ledger", "B"),
    (170, "Other Diretc Expenses", "ledger", "B"),
    (171, "AMC CHARGES EXP - 18%", "ledger", "B"),
    (172, "LOCAL SPARE PARTS EXP", "ledger", "B"),
    (173, "SERVICE CHARGES ACCOUNTS (MACHINERY & PARTS)", "ledger", "B"),
    (174, "SERVICE CHARGES ACCOUNTS (MACHINERY & PARTS)-18%", "ledger", "B"),
    (175, "PACKING EXP", "ledger", "B"),
    (176, "PACKAGING CHARGERS ON IMPORT", "ledger", "B"),
    (177, "PACKING MATERIAL CHAGRES", "ledger", "B"),
    (178, "PACKING MATRIAL ACCOUNTS", "ledger", "B"),
    (179, "PACKING MATRIAL ACCOUNTS 18%", "ledger", "B"),
    (180, "SAMPLE FABRIC EXPENSES", "ledger", "B"),
    (181, "SAMPLE FABRIC PURCHASE EXP.(18%)", "ledger", "B"),
    (182, "SAMPLE FABRIC PURCHASE EXP.(5%)", "ledger", "B"),
    (183, "SAMPLE FABRIC PURCHASE EXP(NON-GST)", "ledger", "B"),
    (184, "SERVICE EXP", "ledger", "B"),
    (185, "SERVICE CHARGE (NON-GST)", "ledger", "B"),
    (186, "SERVICE CHARGES ACCOUNTS (ITC)", "ledger", "B"),
    (187, "INVENTORY OVERHEAD COST", "ledger", "B"),
    (188, "INVENTORY - TRANSFER TO FA", "ledger", "B"),
    (189, "PROVISION FOR IMPORT-FOC(PL)", "ledger", "B"),
    (190, "Gross Profit :", "group_c", "C"),
    (191, "Income Statement:", "group_c", "C"),
    (192, "Indirect Incomes", "group_c", "C"),
    (193, "INTEREST INCOME", "ledger", "B"),
    (194, "INTEREST INCOME ON UNSECURED LOAN A/C(RELATED)", "ledger", "B"),
    (195, "FD INTEREST RECEIVED ACCOUNTS", "ledger", "B"),
    (196, "INTEREST ON FD", "ledger", "B"),
    (197, "OTHER INCOME", "ledger", "B"),
    (198, "PROFIT ON SALE OF ASSET", "ledger", "B"),
    (199, "RENT (INCOME)", "ledger", "B"),
    (200, "RENT (INCOME-RELATED)", "ledger", "B"),
    (201, "", "subtotal_c", "C"),
    (202, "Indirect Expenses", "group_c", "C"),
    (203, "ADVERTISING PROMOTION EXP", "ledger", "B"),
    (204, "ADVERTISING EXP ACCOUNTS @5%", "ledger", "B"),
    (205, "ADVERTISING EXP ACCOUNTS-NON GST", "ledger", "B"),
    (206, "ADVERTISING EXP(ITC) - 18%", "ledger", "B"),
    (207, "EXHIBITION EXP - 18%", "ledger", "B"),
    (208, "EXHIBITION EXP - 5%", "ledger", "B"),
    (209, "EXHIBITION EXP-NON GST", "ledger", "B"),
    (210, "COMMISSION & BROKERAGE EXP", "ledger", "B"),
    (211, "COMMISSION & BROKERAGE ACCOUNTS-E", "ledger", "B"),
    (212, "COMMISSION & BROKERAGE ACCOUNTS (ITC)", "ledger", "B"),
    (213, "COMMISSION & BROKERAGE ACCOUNTS-NON GST", "ledger", "B"),
    (214, "COMPUTER & PRINTER EXP", "ledger", "B"),
    (215, "COMPUTER & PRINTER EXPENSES", "ledger", "B"),
    (216, "COMPUTER & PRINTER EXPENSES 18%", "ledger", "B"),
    (217, "INTERNET & WEBSITE EXPENSES", "ledger", "B"),
    (218, "SOFTWARE CHARGES  ACCOUNTS", "ledger", "B"),
    (219, "TONER REFELING EXP-18%", "ledger", "B"),
    (220, "DEPRECIATION & AMORTISATION EXPENSE", "ledger", "B"),
    (221, "DEPRECIATION", "ledger", "B"),
    (222, "DONATION SUBSCRIPTION", "ledger", "B"),
    (223, "DONATION EXP", "ledger", "B"),
    (224, "DONATION EXP(CSR)", "ledger", "B"),
    (225, "ELECTRICITY EXPENSES", "ledger", "B"),
    (226, "ELECTRICITY EXPENSES ACCOUNTS", "ledger", "B"),
    (227, "EMPLOYEE COST", "ledger", "B"),
    (228, "EMPLOYEE BENEFITS EXPENSES", "ledger", "B"),
    (229, "EMPLOYEE LABOUR WELFARE CHARGES", "ledger", "B"),
    (230, "EMPLOYER ESIC CONTRIBUTION", "ledger", "B"),
    (231, "EMPLOYER LABOUR WELFARE CONTRIBUTION", "ledger", "B"),
    (232, "EMPLOYER PF CONTRIBUTION", "ledger", "B"),
    (233, "EMPLOYER PROFESSIONAL TAX", "ledger", "B"),
    (234, "INTEREST ON PT", "ledger", "B"),
    (235, "PENALTY CHARGES ON PF", "ledger", "B"),
    (236, "PENALTY CHARGES ON PT", "ledger", "B"),
    (237, "PF ADMIN CHARGES", "ledger", "B"),
    (238, "EMPLOYEE SALARY EXPENSES", "ledger", "B"),
    (239, "DRIVER SALARY EXP A/C", "ledger", "B"),
    (240, "EMPLOYEE BONUS ACCOUNTS", "ledger", "B"),
    (241, "EMPLOYEE INCENTIVES EXP ACCOUNTS", "ledger", "B"),
    (242, "EMPLOYEE PERFORMANCE EXPENSES", "ledger", "B"),
    (243, "EMPLOYEE SALARY ACCOUNTS", "ledger", "B"),
    (244, "GRATUITY EXPENSES", "ledger", "B"),
    (245, "MANAGERIAL REMUNERATION", "ledger", "B"),
    (246, "DIRECTORS SALARY EXP ACCOUNTS", "ledger", "B"),
    (247, "STAFF WELFARE EXPENSES", "ledger", "B"),
    (248, "BUSINESS PROMOTION EXPENSE- NON GST", "ledger", "B"),
    (249, "STAFF RECRUITMENT EXPESES-18%", "ledger", "B"),
    (250, "STAFF WELFARE EXP 18%", "ledger", "B"),
    (251, "STAFF WELFARE EXP 5%", "ledger", "B"),
    (252, "STAFF WELFARE EXP ACCOUNTS", "ledger", "B"),
    (253, "EMPLOYEE TRAVEL COST", "ledger", "B"),
    (254, "HOTEL EXPENSES", "ledger", "B"),
    (255, "HOTEL EXP ACCOUNT (12%)", "ledger", "B"),
    (256, "HOTEL EXP ACCOUNT (18%)", "ledger", "B"),
    (257, "HOTEL EXP ACCOUNT (5%)", "ledger", "B"),
    (258, "HOTEL EXP ACCOUNT (INTERNATIONAL)", "ledger", "B"),
    (259, "HOTEL EXP ACCOUNT (NON-GST)", "ledger", "B"),
    (260, "TRAVELLING EXP", "ledger", "B"),
    (261, "TC-D/A", "ledger", "B"),
    (262, "TC-FUEL", "ledger", "B"),
    (263, "TC-HOTEL", "ledger", "B"),
    (264, "TC-MEALS", "ledger", "B"),
    (265, "TC-OTHERS", "ledger", "B"),
    (266, "TC-TRAVEL", "ledger", "B"),
    (267, "TRAVELLING AND CON. EXP-INT.DIRECTOR", "ledger", "B"),
    (268, "TRAVELLING AND CON. EXP-INT.DIRECTOR-18%", "ledger", "B"),
    (269, "TRAVELLING AND CON. EXP-INTERNATIONAL", "ledger", "B"),
    (270, "TRAVELLING AND CON. EXP-INTERNATIONAL-18%", "ledger", "B"),
    (271, "TRAVELLING AND CONVEYANCE (DOLLAR)", "ledger", "B"),
    (272, "TRAVELLING AND CONVEYANCE EXPENSE", "ledger", "B"),
    (273, "TRAVELLING AND CONVEYANCE EXPENSE- 18%", "ledger", "B"),
    (274, "TRAVELLING AND CONVEYANCE EXPENSE- 18%-DIRECTOR", "ledger", "B"),
    (275, "TRAVELLING AND CONVEYANCE EXPENSE- 18% (SERVICE)", "ledger", "B"),
    (276, "TRAVELLING AND CONVEYANCE EXPENSE-DIRECTOR", "ledger", "B"),
    (277, "TRAVELLING EXP - 5%", "ledger", "B"),
    (278, "TRAVELLING EXP - 5%-DIRECTOR", "ledger", "B"),
    (279, "VEHICAL EXP", "ledger", "B"),
    (280, "CAR & VEHICLE EXP ACCOUNTS (NON-GST)", "ledger", "B"),
    (281, "PETROL & DIESEL EXP ACCOUNTS", "ledger", "B"),
    (282, "ENVIRONMENTAL PROTECTION", "ledger", "B"),
    (283, "ENVIRONMENTAL PROTECTION EXPENSES A/C", "ledger", "B"),
    (284, "ENVIRONMENTAL PROTECTION EXPENSES A/C-5%", "ledger", "B"),
    (285, "FINANCE COST", "ledger", "B"),
    (286, "BANK CHARGES-IMPORT DIRECT PAYMENT", "ledger", "B"),
    (287, "BANK CHARGES ACC-IMPORT DIRECT PAYMENT", "ledger", "B"),
    (288, "COMMISSION CHARGES-IMPORT PAYMENT", "ledger", "B"),
    (289, "CURRENCY CONVERSION TAX CGST", "ledger", "B"),
    (290, "CURRENCY CONVERSION TAX SGST", "ledger", "B"),
    (291, "GFID CHARGES-IMPORT", "ledger", "B"),
    (292, "LOAN PROCESSING CHARGES ACCOUNTS-ITC", "ledger", "B"),
    (293, "SWIFT CHARGES-IMPORT", "ledger", "B"),
    (294, "BANK CHARGES-LC-BC", "ledger", "B"),
    (295, "BANK CHARGES-LC/BC", "ledger", "B"),
    (296, "COMMISSION CHARGES-LC-BC PAYMENT", "ledger", "B"),
    (297, "COMMSSION IN LIEU OF EXCHANGE-LC-BC", "ledger", "B"),
    (298, "INTEREST EXP ON LC/BC PAYMENT", "ledger", "B"),
    (299, "SWIFT CHARGES-LC-BC", "ledger", "B"),
    (300, "INTEREST COST", "ledger", "B"),
    (301, "BANK INTEREST ACCOUNTS", "ledger", "B"),
    (302, "INTEREST ON ESIC/PF", "ledger", "B"),
    (303, "INTEREST ON GST PAYMENT", "ledger", "B"),
    (304, "INTEREST ON LOAN (CAR LOAN)", "ledger", "B"),
    (305, "INTEREST ON LOAN -TERM LOAN", "ledger", "B"),
    (306, "INTEREST ON TDS-TCS", "ledger", "B"),
    (307, "INTEREST ON  UNSECURED LOAN A/C", "ledger", "B"),
    (308, "OTHER BANK CHARGES", "ledger", "B"),
    (309, "BANK CHARGES ACCOUNTS", "ledger", "B"),
    (310, "CREDIT CARD CHARGES", "ledger", "B"),
    (311, "LOAN PROCESSING CHARGES ACCOUNTS", "ledger", "B"),
    (312, "INSURANCE EXP", "ledger", "B"),
    (313, "CAR & VEHICLE INSURANCE EXP ACCOUNTS", "ledger", "B"),
    (314, "CAR & VEHICLE INSURANCE EXP ACCOUNTS@18%", "ledger", "B"),
    (315, "INSURANCE EXPENSE ACCOUNT@18%", "ledger", "B"),
    (316, "INSURANCE EXP-NON GST", "ledger", "B"),
    (317, "MARINE INSURANCE EXP - 18%", "ledger", "B"),
    (318, "MARINE INSURANCE EXP - 18%- (PAPER)", "ledger", "B"),
    (319, "MARINE INSURANCE EXP-NON GST", "ledger", "B"),
    (320, "TRAVELLING INSURANCE", "ledger", "B"),
    (321, "LABORATORY EXPENCES", "ledger", "B"),
    (322, "LABORATORY EXPENSE", "ledger", "B"),
    (323, "LABORATORY EXPENSE-18%", "ledger", "B"),
    (324, "LABORATORY EXPENSE-5%", "ledger", "B"),
    (325, "Learning and Development", "ledger", "B"),
    (326, "TRAINING EXPENSES-18%", "ledger", "B"),
    (327, "LEGAL AND PROFESSIONAL EXP", "ledger", "B"),
    (328, "CONSULATANCY CHARGE", "ledger", "B"),
    (329, "CONSULTANCY CHARGES-NON GST", "ledger", "B"),
    (330, "CONSULTANCY CHARGES-RCM", "ledger", "B"),
    (331, "LEGAL AND PROFESSIONAL A/C(ITC)", "ledger", "B"),
    (332, "LEGAL FEES(RCM)", "ledger", "B"),
    (333, "PROFESSIONAL FEES A/C@18%", "ledger", "B"),
    (334, "TECHNICAL & CONSULTANCY FEES (ITC)", "ledger", "B"),
    (335, "TRADEMARK COUNSULTANCY (ITC)", "ledger", "B"),
    (336, "OFFICE EXPENSES", "ledger", "B"),
    (337, "OTHER EXPENSES", "ledger", "B"),
    (338, "ROC TAXES", "ledger", "B"),
    (339, "OFFICE CONSUMABLES EXP", "ledger", "B"),
    (340, "OFFICE EXP ACCOUNTS", "ledger", "B"),
    (341, "OFFICE EXPENSES 18%", "ledger", "B"),
    (342, "OFFICE EXPENSES 5%", "ledger", "B"),
    (343, "OFFICE EXPENSES @12%", "ledger", "B"),
    (344, "OFFICE MAINTANCE EXP ACCOUNT", "ledger", "B"),
    (345, "OFFICE TITAANIUM AREA REPARING", "ledger", "B"),
    (346, "ROUND OFF", "ledger", "B"),
    (347, "WATER CHARGES", "ledger", "B"),
    (348, "PRINTING AND COURIER", "ledger", "B"),
    (349, "COURIER & POSTAGE EXP", "ledger", "B"),
    (350, "POSTAGE AND COURIER EXPENSES - 18%", "ledger", "B"),
    (351, "POSTAGE AND COURIER EXPENSES -HEAD", "ledger", "B"),
    (352, "POSTAGE AND COURIER EXPENSES-(NON GST)", "ledger", "B"),
    (353, "POSTAGE AND COURIER EXPENSES -SPARE", "ledger", "B"),
    (354, "POSTAGE & COURIER EXP (NON-GST)", "ledger", "B"),
    (355, "PRINTING STATIONERY EXP", "ledger", "B"),
    (356, "PRINTING & STATIONERY @12%", "ledger", "B"),
    (357, "PRINTING & STATIONERY @18%", "ledger", "B"),
    (358, "PRINTING & STATIONERY EXP ACCOUNTS", "ledger", "B"),
    (359, "RENT EXPENSES", "ledger", "B"),
    (360, "GODOWN RENT", "ledger", "B"),
    (361, "GODOWN RENT-RCM", "ledger", "B"),
    (362, "RENT-TITAANIUM", "ledger", "B"),
    (363, "REPAIR & MAINTENANCE COST", "ledger", "B"),
    (364, "ELECTRICAL EXPENSES", "ledger", "B"),
    (365, "ELECTRICAL EXPENSES  - 18%", "ledger", "B"),
    (366, "ELECTRICAL EXPENSES ACCOUNTS", "ledger", "B"),
    (367, "REPAIR & MAINTENANCE EXP (OFFICE)", "ledger", "B"),
    (368, "FIT OUT CHARGES (MUMBAI OFFICE)", "ledger", "B"),
    (369, "HOJIWALA GODOWN  REPAIRS", "ledger", "B"),
    (370, "HOJIWALA GODOWN REPAIRS-12%", "ledger", "B"),
    (371, "HOJIWALA GODOWN REPAIRS-18%", "ledger", "B"),
    (372, "HVAC CHARGES (MUMBAI OFFICE)", "ledger", "B"),
    (373, "MAINTENANCE EXP-TITAANIUM", "ledger", "B"),
    (374, "MUMBAI OFFICE MAINTENANCE ACCOUNTS", "ledger", "B"),
    (375, "OFFICE TITAANIUM AREA REPARING-18%", "ledger", "B"),
    (376, "REPAIR & MAINTENANCE CHARGES", "ledger", "B"),
    (377, "REPAIR & MAINTENANCE CHARGES@18%", "ledger", "B"),
    (378, "REPAIR & MAINTENANCE CHARGES@28%", "ledger", "B"),
    (379, "SACHIN GODOWN REPAIRS-18%", "ledger", "B"),
    (380, "SECURITY EXP", "ledger", "B"),
    (381, "SECURITY EXP", "ledger", "B"),
    (382, "SECURITY EXP(RCM)", "ledger", "B"),
    (383, "SUBSCRIPTION & MEMBERSHIP EXP", "ledger", "B"),
    (384, "AI SUBSCRIPTION", "ledger", "B"),
    (385, "AI SUBSCRIPTION EXPENSE (RCM)", "ledger", "B"),
    (386, "BUSINESS PROMOTION EXPENSE-18%", "ledger", "B"),
    (387, "REGISTRATION CHARGES", "ledger", "B"),
    (388, "SUBSCRIPTION EXPENSE (RCM)", "ledger", "B"),
    (389, "SUBSCRIPTION & MEMBERSHIP EXP (ITC)", "ledger", "B"),
    (390, "SUBSCRIPTION & MEMBERSHIP EXP (NON-GST)", "ledger", "B"),
    (391, "TAXES EXP", "ledger", "B"),
    (392, "GST EXPENSE", "ledger", "B"),
    (393, "GST LATE FEES", "ledger", "B"),
    (394, "GUJARAT POLLUTION CONTROL BOARD", "ledger", "B"),
    (395, "PROPERTY TAX", "ledger", "B"),
    (396, "STAMP PURCHASE AND NOTERY EXP ACCOUNTS", "ledger", "B"),
    (397, "TELEPHONE EXP", "ledger", "B"),
    (398, "TELEPHONE & MOBILE EXPENSES - 18%", "ledger", "B"),
    (399, "TELEPHONE & MOBILE EXPENSES(NON-GST)", "ledger", "B"),
    (400, "TRANSPORTATION EXP", "ledger", "B"),
    (401, "CARRIAGE OUTWARD ACCOUNTS", "ledger", "B"),
    (402, "CARRIAGE OUTWARD ACCOUNTS - 12%", "ledger", "B"),
    (403, "CARRIAGE OUTWARD ACCOUNTS (RCM)", "ledger", "B"),
    (404, "CARRIAGE OUTWARD EXP ACCOUNTS@5%", "ledger", "B"),
    (405, "CARRIRAGE OUTWARD ACCOUNTS- 18%", "ledger", "B"),
    (406, "LOADING & UNLOADING EXP", "ledger", "B"),
    (407, "LOADING & UNLOADING ACCOUNTS@18%", "ledger", "B"),
    (408, "LOADING & UNLOADING CHARGES (RCM)", "ledger", "B"),
    (409, "LOADING & UNLOADING EXP-(NON RCM)", "ledger", "B"),
    (410, "LOADING & UNLOADING EXP (PAPER)", "ledger", "B"),
    (411, "Excess Short Income Tax 24-25", "ledger", "B"),
    (412, "Nett Profit:", "group_c", "C"),
]

IDX = {row[0]: i for i, row in enumerate(ACCOUNT_ORDER)}


def post_to_tally(xml_body: str, host: str = TALLY_HOST, port: int = TALLY_PORT) -> str:
    url = f"http://{host}:{port}"
    encoded = xml_body.encode("utf-8")
    req = urllib.request.Request(
        url,
        data=encoded,
        headers={"Content-Type": "text/xml; charset=utf-8"},
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.read().decode("utf-8", errors="replace")
    except urllib.error.URLError as e:
        raise RuntimeError(
            f"Cannot connect to Tally at {url}.\n"
            f"  Cause: {e}\n"
            f"  Make sure Tally is running, the HTTP XML server is enabled,\n"
            f"  and at least one company is open."
        ) from e


def get_open_companies(host: str = TALLY_HOST, port: int = TALLY_PORT) -> list:
    xml = """<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>AllCompanies</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION ISMODIFY="No" NAME="AllCompanies">
            <TYPE>Company</TYPE>
            <FETCH>Name,StartingFrom,EndingAt</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""
    resp = post_to_tally(xml, host, port)
    companies = []
    try:
        root = ET.fromstring(resp)
        for elem in root.iter("COMPANY"):
            name_elem = elem.find("NAME")
            if name_elem is not None and name_elem.text and name_elem.text.strip():
                name = name_elem.text.strip()
                if name not in companies:
                    companies.append(name)
    except ET.ParseError:
        pass

    if not companies:
        import re
        found = re.findall(r'<NAME[^>]*>(.*?)</NAME>', resp, re.DOTALL | re.IGNORECASE)
        for m in found:
            name = m.strip()
            if name and name not in companies:
                companies.append(name)

    if not companies:
        print("WARNING: Could not detect open companies from Tally.", file=sys.stderr)
        print(f"Raw Tally response (first 800 chars): {resp[:800]}", file=sys.stderr)
    return companies


def parse_tally_amount(element) -> float:
    if element is None:
        return 0.0
    text = (element.text or "").strip()
    if not text:
        return 0.0
    try:
        amount = float(text)
    except ValueError:
        return 0.0
    type_attr = (element.get("TYPE") or "").upper()
    return amount if type_attr != "DR" else -amount


def fetch_ledger_balances(period_start, period_end, company: str, host: str = TALLY_HOST, port: int = TALLY_PORT) -> dict:
    from_str = period_start.strftime("%Y%m%d")
    to_str = period_end.strftime("%Y%m%d")
    xml = f"""<ENVELOPE>
  <HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
  <BODY>
    <EXPORTDATA>
      <REQUESTDESC>
        <STATICVARIABLES>
          <SVFROMDATE>{from_str}</SVFROMDATE>
          <SVTODATE>{to_str}</SVTODATE>
          <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
          <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
        </STATICVARIABLES>
        <FORMULACOLLECTION>
          <COLLECTION NAME="PL_LEDGERS">
            <TYPE>Ledger</TYPE>
            <FETCH>Name,Parent,ClosingBalance</FETCH>
            <FILTER>IsPLAffecting</FILTER>
          </COLLECTION>
        </FORMULACOLLECTION>
      </REQUESTDESC>
    </EXPORTDATA>
  </BODY>
</ENVELOPE>"""
    resp = post_to_tally(xml, host, port)
    result = {}
    try:
        root = ET.fromstring(resp)
        for ledger in root.iter("LEDGER"):
            name_el = ledger.find("NAME")
            bal_el = ledger.find("CLOSINGBALANCE")
            if name_el is not None and name_el.text:
                name = name_el.text.strip().upper()
                result[name] = parse_tally_amount(bal_el)
    except ET.ParseError as e:
        print(f"  XML parse error for ledger response: {e}", file=sys.stderr)
    return result


def fetch_group_balances(period_start, period_end, company: str, host: str = TALLY_HOST, port: int = TALLY_PORT) -> dict:
    from_str = period_start.strftime("%Y%m%d")
    to_str = period_end.strftime("%Y%m%d")
    xml = f"""<ENVELOPE>
  <HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
  <BODY>
    <EXPORTDATA>
      <REQUESTDESC>
        <STATICVARIABLES>
          <SVFROMDATE>{from_str}</SVFROMDATE>
          <SVTODATE>{to_str}</SVTODATE>
          <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
          <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
        </STATICVARIABLES>
        <FORMULACOLLECTION>
          <COLLECTION NAME="PL_GROUPS">
            <TYPE>Group</TYPE>
            <FETCH>Name,ClosingBalance</FETCH>
            <FILTER>IsPLAffecting</FILTER>
          </COLLECTION>
        </FORMULACOLLECTION>
      </REQUESTDESC>
    </EXPORTDATA>
  </BODY>
</ENVELOPE>"""
    resp = post_to_tally(xml, host, port)
    result = {}
    try:
        root = ET.fromstring(resp)
        for grp in root.iter("GROUP"):
            name_el = grp.find("NAME")
            bal_el = grp.find("CLOSINGBALANCE")
            if name_el is not None and name_el.text:
                name = name_el.text.strip().upper()
                result[name] = parse_tally_amount(bal_el)
    except ET.ParseError as e:
        print(f"  XML parse error for group response: {e}", file=sys.stderr)
    return result


def parse_date(s: str) -> date:
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Cannot parse date: {s!r}. Use DD-MM-YYYY or YYYY-MM-DD.")


def get_months(from_date: date, to_date: date) -> list:
    months = []
    y, m = from_date.year, from_date.month
    while True:
        last_day = calendar.monthrange(y, m)[1]
        start = date(y, m, 1)
        end = date(y, m, last_day)
        months.append((start, end))
        if (y, m) >= (to_date.year, to_date.month):
            break
        m += 1
        if m > 12:
            m = 1
            y += 1
    return months


def month_label(start: date, end: date) -> str:
    return f"1-{start.strftime('%b-%y')} to {end.day}-{end.strftime('%b-%y')}"


def lookup(name: str, ledger_data: dict, group_data: dict) -> float:
    key = name.strip().upper()
    if key in ledger_data:
        return ledger_data[key]
    if key in group_data:
        return group_data[key]
    for prefix in ("ADD: ", "LESS: ", "ADD:", "LESS:"):
        stripped = key.removeprefix(prefix).strip()
        if stripped in ledger_data:
            return ledger_data[stripped]
        if stripped in group_data:
            return group_data[stripped]
    return 0.0


def compute_amounts_for_month(ledger_data: dict, group_data: dict) -> list:
    amounts = [0.0] * len(ACCOUNT_ORDER)

    for i, (row_num, name, rtype, col) in enumerate(ACCOUNT_ORDER):
        if rtype in ("ledger", "group_c", "group_b"):
            amounts[i] = lookup(name, ledger_data, group_data)

    i77 = IDX.get(77); i13 = IDX.get(13); i62 = IDX.get(62)
    if i77 is not None and i13 is not None and i62 is not None:
        amounts[i77] = amounts[i13] + amounts[i62]

    i82 = IDX.get(82); i79 = IDX.get(79); i80 = IDX.get(80); i81 = IDX.get(81)
    if all(x is not None for x in [i82, i79, i80, i81]):
        amounts[i82] = amounts[i79] + amounts[i80] - amounts[i81]

    i78 = IDX.get(78); i83 = IDX.get(83)
    if i78 is not None and amounts[i78] == 0.0 and i82 is not None and i83 is not None:
        amounts[i78] = amounts[i82] + abs(amounts[i83])

    i190 = IDX.get(190)
    if i190 is not None and amounts[i190] == 0.0 and i77 is not None and i78 is not None:
        amounts[i190] = amounts[i77] - amounts[i78]

    i201 = IDX.get(201); i192 = IDX.get(192)
    if i201 is not None and i190 is not None and i192 is not None:
        amounts[i201] = amounts[i190] + amounts[i192]

    i412 = IDX.get(412); i202 = IDX.get(202); i411 = IDX.get(411)
    if i412 is not None and amounts[i412] == 0.0 and i201 is not None and i202 is not None:
        excess_tax = amounts[i411] if i411 is not None else 0.0
        amounts[i412] = amounts[i201] - abs(amounts[i202]) - excess_tax

    return amounts


BOLD_FONT = Font(bold=True, name="Calibri", size=10)
NORMAL_FONT = Font(bold=False, name="Calibri", size=10)
NUM_FORMAT = "#,##0.00"
COL_A_WIDTH = 61.56
DATA_COL_WIDTH = 11.67
BOLD_RTYPES = {"group_c", "group_b", "subtotal_c", "subtotal_b"}


def write_company_sheet(ws, company: str, months: list, monthly_amounts: list):
    n = len(months)
    ws.column_dimensions["A"].width = COL_A_WIDTH
    for m_idx in range(n):
        ws.column_dimensions[get_column_letter(2 + m_idx * 2)].width = DATA_COL_WIDTH
        ws.column_dimensions[get_column_letter(3 + m_idx * 2)].width = DATA_COL_WIDTH

    fy_label = f"{months[0][0].strftime('%d-%m-%y')}TO{months[-1][1].strftime('%d-%m-%y')}"
    header_lines = [
        f"{company} ({fy_label})",
        "7111/A2,ROAD NO.71, G.I.D.C,SACHIN SURAT - 394230",
        "DELIVERY ADD:PLOT NO.C-15/3, HOJIWALA INDUSTRIAL ESTATE, ROAD NO.14,",
        "HOJIWALA SURAT-394230",
        "GST NUMBER :24AAHCM0016A1ZD",
        "Email - Accounts@orangeotec.Com",
        "Co.No.85110 94943",
        "MSME:UDYAM-GJ-22-0028640-MEDIUM",
        "Profit & Loss A/c",
    ]
    for row_i, text in enumerate(header_lines, start=1):
        ws.cell(row_i, 1, value=text).font = BOLD_FONT if row_i in (1, 9) else NORMAL_FONT

    for m_idx in range(n):
        cell = ws.cell(10, 2 + m_idx * 2, value=company)
        cell.font = NORMAL_FONT
        cell.alignment = Alignment(horizontal="center")

    ws.cell(11, 1, value="Particulars").font = BOLD_FONT
    for m_idx, (mstart, mend) in enumerate(months):
        cell = ws.cell(11, 2 + m_idx * 2, value=month_label(mstart, mend))
        cell.font = BOLD_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    ws.freeze_panes = "B12"

    for i, (row_num, name, rtype, col) in enumerate(ACCOUNT_ORDER):
        is_bold = rtype in BOLD_RTYPES
        ws.cell(row_num, 1, value=name if name else None).font = BOLD_FONT if is_bold else NORMAL_FONT
        for m_idx in range(n):
            amount = monthly_amounts[m_idx][i]
            if abs(amount) < 0.005:
                continue
            data_col = (3 if col == "C" else 2) + m_idx * 2
            cell = ws.cell(row_num, data_col, value=amount)
            cell.number_format = NUM_FORMAT
            cell.font = BOLD_FONT if is_bold else NORMAL_FONT


def write_excel(companies_data: dict, months: list, output_path: str):
    wb = openpyxl.Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    for company, monthly_amounts in companies_data.items():
        write_company_sheet(wb.create_sheet(title=company[:31]), company, months, monthly_amounts)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(output_path)
    except PermissionError:
        raise RuntimeError(f"Cannot save to {output_path!r} — file may be open in Excel.\nClose the file and retry.")


_PERIOD_COL_MAP = {"2627": "26-27", "2526": "25-26", "2425": "24-25"}


def _get_period_companies(period_code: str) -> list:
    """Return company names for the given period code from TALLY Company_List.xlsx."""
    col_name = _PERIOD_COL_MAP.get(str(period_code))
    if not col_name:
        raise ValueError(f"Unknown period code '{period_code}'. Use 2627, 2526, or 2425.")
    candidates = [
        Path.cwd() / "Files" / "TALLY Company_List.xlsx",
        Path(__file__).parent.parent.parent.parent / "Files" / "TALLY Company_List.xlsx",
    ]
    co_list_path = next((p for p in candidates if p.exists()), None)
    if co_list_path is None:
        raise FileNotFoundError(
            "Cannot find Files/TALLY Company_List.xlsx. Run from the project root."
        )
    import openpyxl as _opxl
    wb2 = _opxl.load_workbook(str(co_list_path), read_only=True)
    ws2 = wb2.active
    headers = None
    col_idx = None
    result = []
    for row in ws2.iter_rows(values_only=True):
        if headers is None:
            if row[0] == "Company Name":
                headers = list(row)
                col_idx = headers.index(col_name)
            continue
        company_name = row[0]
        comment = row[3] if len(row) > 3 else None
        if not company_name or comment == "Not To USE":
            continue
        val = row[col_idx] if col_idx is not None and len(row) > col_idx else None
        if val == 1:
            result.append(str(company_name).strip())
    wb2.close()
    return result


def parse_args():
    p = argparse.ArgumentParser(description="Fetch P&L from Tally for all open companies")
    p.add_argument("--from-date", default="01-04-2025", help="Start date DD-MM-YYYY")
    p.add_argument("--to-date", default="31-03-2026", help="End date DD-MM-YYYY")
    p.add_argument("--output", default="Files/PandL_Output.xlsx", help="Output Excel path")
    p.add_argument("--port", type=int, default=TALLY_PORT, help="Tally XML server port")
    p.add_argument("--period", default=None,
                   help="Period code 2627/2526/2425 — filters companies via TALLY Company_List.xlsx")
    return p.parse_args()


def main():
    args = parse_args()
    from_date = parse_date(args.from_date)
    to_date = parse_date(args.to_date)
    months = get_months(from_date, to_date)
    print(f"Date range: {from_date} to {to_date} ({len(months)} months)")

    print(f"Connecting to Tally on port {args.port}...")
    companies = get_open_companies(host=TALLY_HOST, port=args.port)
    if not companies:
        print("ERROR: No companies detected in Tally. Open at least one company and retry.")
        sys.exit(1)

    if args.period:
        whitelist = _get_period_companies(args.period)
        whitelist_upper = {w.upper() for w in whitelist}
        companies = [c for c in companies if c.upper() in whitelist_upper]
        print(f"Period filter {args.period}: keeping {len(companies)} company/companies.")
        if not companies:
            print("ERROR: No matching companies found for this period. Exiting.")
            sys.exit(1)

    print(f"Found {len(companies)} company/companies: {', '.join(companies)}")

    companies_data = {}
    for company in companies:
        print(f"\nProcessing: {company}")
        monthly_amounts = []
        for mstart, mend in months:
            print(f"  Fetching {month_label(mstart, mend)}...", end="", flush=True)
            ledger_data = fetch_ledger_balances(mstart, mend, company, host=TALLY_HOST, port=args.port)
            group_data = fetch_group_balances(mstart, mend, company, host=TALLY_HOST, port=args.port)
            amounts = compute_amounts_for_month(ledger_data, group_data)
            monthly_amounts.append(amounts)
            net_idx = IDX.get(412)
            print(f" Net Profit: {amounts[net_idx]:,.2f}" if net_idx is not None else "")
        companies_data[company] = monthly_amounts

    print(f"\nWriting Excel to: {args.output}")
    write_excel(companies_data, months, args.output)

    print("\n=== SUMMARY ===")
    for company, monthly_amounts in companies_data.items():
        total_net = sum(amounts[IDX[412]] for amounts in monthly_amounts if IDX.get(412) is not None)
        print(f"  {company}: Total Net Profit = {total_net:,.2f}")
    print(f"\nFile saved: {args.output}")


if __name__ == "__main__":
    main()
