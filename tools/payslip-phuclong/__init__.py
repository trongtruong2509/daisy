"""
Payslip automation tool for Phuc Long company.

Reads employee data from Excel, generates password-protected PDF payslips,
and sends them via Outlook email.

Usage:
    cd tools/payslip-phuclong
    python main.py
"""

from config import load_payslip_config, PayslipConfig
from excel_reader import ExcelReader
from validator import DataValidator
from payslip_generator import PayslipGenerator
from pdf_converter import PdfConverter
from email_composer import EmailComposer

__all__ = [
    "load_payslip_config",
    "PayslipConfig",
    "ExcelReader",
    "DataValidator",
    "PayslipGenerator",
    "PdfConverter",
    "EmailComposer",
]
