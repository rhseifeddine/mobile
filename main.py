import hashlib
import json
import logging
import math
import os
import random
import re
import shutil
import sqlite3
import sys
import threading
import time
import traceback
import urllib.parse
import webbrowser
import zipfile
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import arabic_reshaper
import openpyxl
from bidi.algorithm import get_display
from fpdf import FPDF, XPos, YPos
from PIL import Image, ImageDraw, ImageFont

# ==============================================================================
# 1. ENVIRONNEMENT, JOURNAUX ET PONTS MATÉRIELS
# ==============================================================================

DEBUG = False
if DEBUG:
    os.environ['KIVY_LOG_LEVEL'] = 'info'
    os.environ['KIVY_NO_CONSOLELOG'] = '0'
else:
    os.environ['KIVY_LOG_LEVEL'] = 'error'
    os.environ['KIVY_NO_CONSOLELOG'] = '1'

from kivy.animation import Animation
from kivy.clock import Clock, mainthread
from kivy.config import Config
from kivy.core.clipboard import Clipboard
from kivy.core.text import LabelBase
from kivy.core.window import Window
from kivy.graphics.context_instructions import PopMatrix, PushMatrix, Rotate
from kivy.lang import Builder
from kivy.metrics import dp
from kivy.properties import (
    BooleanProperty,
    ColorProperty,
    ListProperty,
    ObjectProperty,
    StringProperty,
)
from kivy.uix.camera import Camera
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.modalview import ModalView
from kivy.uix.recycleview import RecycleView
from kivy.uix.recycleview.views import RecycleDataViewBehavior
from kivy.uix.spinner import Spinner
from kivy.utils import platform
from kivymd import fonts_path
from kivymd.app import MDApp
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.button import (
    MDFlatButton,
    MDIconButton,
    MDRaisedButton,
)
from kivymd.uix.card import MDCard, MDSeparator
from kivymd.uix.dialog import MDDialog
from kivymd.uix.fitimage import FitImage
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.gridlayout import MDGridLayout
from kivymd.uix.label import MDIcon, MDLabel
from kivymd.uix.list import (
    ILeftBody,
    IRightBodyTouch,
    MDList,
)
from kivymd.uix.pickers import MDDatePicker
from kivymd.uix.screen import MDScreen
from kivymd.uix.screenmanager import MDScreenManager
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.selectioncontrol import MDCheckbox
from kivymd.uix.textfield import MDTextField

if DEBUG:
    Config.set('kivy', 'log_level', 'info')
    Config.set('kivy', 'log_enable', 1)
else:
    Config.set('kivy', 'log_level', 'error')
    Config.set('kivy', 'log_enable', 0)
Config.write()

try:
    from pyzbar.pyzbar import decode as zbar_decode
except ImportError:
    zbar_decode = None

BluetoothAdapter = None
UUID = None
ToneGenerator = None
PythonActivity = None
if platform == 'android':
    try:
        from jnius import autoclass

        BluetoothAdapter = autoclass('android.bluetooth.BluetoothAdapter')
        UUID = autoclass('java.util.UUID')
        ToneGenerator = autoclass('android.media.ToneGenerator')
        PythonActivity = autoclass('org.kivy.android.PythonActivity')
    except Exception as exc:
        if DEBUG:
            print(f'[ERROR] Android JNI init: {exc}')

APP_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_FILE = os.path.join(APP_DIR, 'font.ttf')
_custom_font_loaded = False
try:
    if os.path.exists(FONT_FILE) and os.path.isfile(FONT_FILE):
        for _fname in ('ArabicFont', 'Roboto', 'RobotoMedium', 'RobotoBold'):
            LabelBase.register(name=_fname, fn_regular=FONT_FILE, fn_bold=FONT_FILE)
        _custom_font_loaded = True
except Exception:
    _custom_font_loaded = False

if not _custom_font_loaded:
    _fb_reg = os.path.join(fonts_path, 'Roboto-Regular.ttf')
    _fb_bold = os.path.join(fonts_path, 'Roboto-Bold.ttf')
    try:
        LabelBase.register(name='ArabicFont', fn_regular=_fb_reg, fn_bold=_fb_bold)
    except Exception:
        LabelBase.register(name='ArabicFont', fn_regular=None, fn_bold=None)

_DEFAULT_RESHAPER = arabic_reshaper.ArabicReshaper(
    configuration={
        'delete_harakat': True,
        'support_ligatures': True,
        'use_unshaped_instead_of_isolated': True,
    }
)
_INPUT_RESHAPER = arabic_reshaper.ArabicReshaper(
    configuration={
        'delete_harakat': True,
        'support_ligatures': False,
        'use_unshaped_instead_of_isolated': True,
    }
)
_AR_ANTI_LIGATURE = '\u200c'
_AR_LAM_ALEF_RE = re.compile('([لﻝﻞﻟﻠ])([آأإاٱﭐﭑﺁﺂﺃﺄﺇﺈﺍﺎ])')
_AR_RAW_RE = re.compile('[\u0600-\u06ff\u0750-\u077f\u08a0-\u08ff]')


def ar_visual(text, shaper=None) -> str:
    """Met en forme le texte arabe pour le rendu visuel."""
    if text is None:
        return ''
    if not isinstance(text, str):
        text = str(text)
    if not text or not _AR_RAW_RE.search(text):
        return text
    try:
        visual = get_display((shaper or _DEFAULT_RESHAPER).reshape(text))
        return _AR_LAM_ALEF_RE.sub(lambda m: m.group(1) + _AR_ANTI_LIGATURE + m.group(2), visual)
    except Exception:
        return text


# ==============================================================================
# 2. CONSTANTES UNIQUES (100% FRANÇAIS, PARITÉ MAGPRO DESKTOP)
# ==============================================================================
class AppConstants:
    APP_NAME = 'MagPro Mobile'
    APP_VERSION = '8.1.0'
    COPYRIGHT_YEAR = '2027'
    DB_NAME = 'magpro_local.db'
    FONT_ARABIC = 'ArabicFont'
    ADMIN_USERNAME = 'ADMIN'
    CURRENCY = 'DA'

    UNLIMITED_STOCK_VALUE = -999999999999.0
    UNLIMITED_STOCK_THRESHOLD = -900000000000.0
    MAX_AMOUNT_LIMIT = Decimal('1000000000')
    MAX_QTY_LIMIT = Decimal('100000000')
    MAX_STOCK_DB_LIMIT = Decimal('9999999999999')
    MAX_PERCENT_LIMIT = Decimal('100')

    DEFAULT_CLIENT_ID = 1
    DEFAULT_SUPPLIER_ID = 2
    DEFAULT_PRODUCT_ID = 1
    DEFAULT_CLIENT_NAME = 'Comptoir'
    DEFAULT_SUPPLIER_NAME = 'Fournisseur'
    DEFAULT_PRODUCT_NAME = 'Autre Article'

    ENTITY_CLIENT = 'account'
    ENTITY_SUPPLIER = 'supplier'

    LOC_STORE = 'store'
    LOC_WAREHOUSE = 'warehouse'
    LOC_LABELS = {
        LOC_STORE: 'Magasin',
        LOC_WAREHOUSE: 'Dépôt',
    }

    PRICE_CAT_DETAIL = 'Détail'
    PRICE_CAT_SEMI = 'Demi-Gros'
    PRICE_CAT_GROS = 'Gros'
    PRICE_CATEGORIES = (PRICE_CAT_DETAIL, PRICE_CAT_SEMI, PRICE_CAT_GROS)

    FAM_ALL = 'TOUS'

    DOC_BV = 'BV'
    DOC_BA = 'BA'
    DOC_FC = 'FC'
    DOC_FF = 'FF'
    DOC_FP = 'FP'
    DOC_DP = 'DP'
    DOC_RC = 'RC'
    DOC_RF = 'RF'
    DOC_BE = 'BE'
    DOC_BS = 'BS'
    DOC_TR = 'TR'
    DOC_BI = 'BI'
    DOC_CLIENT_PAY = 'CLIENT_PAY'
    DOC_SUPPLIER_PAY = 'SUPPLIER_PAY'

    SALE_DOC_TYPES = (DOC_BV, DOC_FC, DOC_FP, DOC_RC, DOC_BS)
    PURCHASE_DOC_TYPES = (DOC_BA, DOC_FF, DOC_DP, DOC_RF, DOC_BI, DOC_BE)
    RETURN_DOC_TYPES = (DOC_RC, DOC_RF)
    NON_FINANCIAL_DOC_TYPES = (DOC_FP, DOC_DP, DOC_BE, DOC_BS, DOC_TR)
    INVOICE_TAX_DOC_TYPES = (DOC_FC, DOC_FF, DOC_FP)
    STAMP_DUTY_DOC_TYPES = (DOC_FC, DOC_FF)
    QTY_ONLY_DOC_TYPES = (DOC_DP, DOC_BE, DOC_BS, DOC_TR)
    YEARLY_RESET_SEQUENCES = (DOC_FC, DOC_FF, DOC_FP, DOC_DP)

    STOCK_MOVEMENTS = {
        DOC_BA: 1,
        DOC_FF: 1,
        DOC_BI: 1,
        DOC_RC: 1,
        DOC_BE: 1,
        DOC_BV: -1,
        DOC_FC: -1,
        DOC_RF: -1,
        DOC_BS: -1,
        DOC_FP: 0,
        DOC_DP: 0,
        DOC_TR: 0,
        DOC_CLIENT_PAY: 0,
        DOC_SUPPLIER_PAY: 0,
    }

    FINANCIAL_FACTORS = {
        DOC_BV: 1,
        DOC_FC: 1,
        DOC_BA: 1,
        DOC_FF: 1,
        DOC_BI: 1,
        DOC_RC: -1,
        DOC_RF: -1,
        DOC_CLIENT_PAY: -1,
        DOC_SUPPLIER_PAY: -1,
        DOC_FP: 0,
        DOC_DP: 0,
        DOC_BE: 0,
        DOC_BS: 0,
        DOC_TR: 0,
    }

    DOC_VISUALS = {
        DOC_BV: {'name': 'Bon de Vente', 'short': 'Vente', 'icon': 'cart', 'color': (0, 0.5, 0.8, 1), 'bg': (0.98, 0.99, 1, 1)},
        DOC_BA: {'name': "Bon d'Achat", 'short': 'Achat', 'icon': 'truck', 'color': (0.9, 0.5, 0, 1), 'bg': (1, 0.98, 0.94, 1)},
        DOC_FC: {'name': 'Facture', 'short': 'Facture', 'icon': 'file-document', 'color': (0, 0.25, 0.8, 1), 'bg': (0.96, 0.97, 1, 1)},
        DOC_FF: {'name': "Facture d'Achat", 'short': 'Fact. Achat', 'icon': 'file-document-edit', 'color': (0.9, 0.3, 0, 1), 'bg': (1, 0.95, 0.92, 1)},
        DOC_FP: {'name': 'Facture Proforma', 'short': 'Proforma', 'icon': 'file-eye', 'color': (0.5, 0, 0.5, 1), 'bg': (0.98, 0.95, 1, 1)},
        DOC_DP: {'name': 'Bon de Commande', 'short': 'Commande', 'icon': 'clipboard-list', 'color': (0, 0.55, 0.55, 1), 'bg': (0.94, 0.99, 0.99, 1)},
        DOC_RC: {'name': 'Avoir sur Vente', 'short': 'Retour Cl.', 'icon': 'keyboard-return', 'color': (0.8, 0, 0, 1), 'bg': (1, 0.95, 0.95, 1)},
        DOC_RF: {'name': 'Avoir sur Achat', 'short': 'Retour Fr.', 'icon': 'undo', 'color': (0, 0.6, 0.6, 1), 'bg': (0.95, 1, 1, 1)},
        DOC_BE: {'name': "Bon d'Entrée", 'short': 'Entrée', 'icon': 'tray-arrow-down', 'color': (0, 0.6, 0.2, 1), 'bg': (0.93, 1, 0.94, 1)},
        DOC_BS: {'name': 'Bon de Sortie', 'short': 'Sortie', 'icon': 'tray-arrow-up', 'color': (0.55, 0.3, 0.1, 1), 'bg': (0.98, 0.95, 0.92, 1)},
        DOC_TR: {'name': 'Bon de Transfert', 'short': 'Transfert', 'icon': 'swap-horizontal', 'color': (0.5, 0, 0.5, 1), 'bg': (0.97, 0.93, 1, 1)},
        DOC_BI: {'name': 'Bon Initial', 'short': 'Stock Init.', 'icon': 'database-plus', 'color': (0, 0.5, 0.5, 1), 'bg': (0.95, 0.99, 0.99, 1)},
        DOC_CLIENT_PAY: {'name': 'Versement', 'short': 'Versement', 'icon': 'cash-plus', 'color': (0, 0.65, 0.2, 1), 'bg': (0.93, 1, 0.94, 1)},
        DOC_SUPPLIER_PAY: {'name': 'Règlement', 'short': 'Règlement', 'icon': 'cash-minus', 'color': (0.85, 0.4, 0, 1), 'bg': (1, 0.95, 0.9, 1)},
    }

    MODE_TO_DOC = {
        'sale': DOC_BV,
        'purchase': DOC_BA,
        'invoice_sale': DOC_FC,
        'invoice_purchase': DOC_FF,
        'proforma': DOC_FP,
        'order_purchase': DOC_DP,
        'return_sale': DOC_RC,
        'return_purchase': DOC_RF,
        'stock_in': DOC_BE,
        'stock_out': DOC_BS,
        'transfer': DOC_TR,
        'initial_stock': DOC_BI,
        'client_payment': DOC_CLIENT_PAY,
        'supplier_payment': DOC_SUPPLIER_PAY,
    }

    MODE_TITLES = {
        'sale': 'Vente (BV)',
        'purchase': 'Achat (BA)',
        'invoice_sale': 'Facture (FC)',
        'invoice_purchase': 'Fact. Achat (FF)',
        'proforma': 'Proforma (FP)',
        'order_purchase': 'Commande (DP)',
        'return_sale': 'Retour Client (RC)',
        'return_purchase': 'Retour Fourn. (RF)',
        'stock_in': "Bon d'Entrée (BE)",
        'stock_out': 'Bon de Sortie (BS)',
        'transfer': 'Transfert (TR)',
        'initial_stock': 'Stock Initial (BI)',
        'manage_products': 'Produits',
        'client_payment': 'Versement',
        'supplier_payment': 'Règlement',
    }

    MODE_COLORS = {
        'sale': 'Green',
        'purchase': 'Orange',
        'invoice_sale': 'Blue',
        'invoice_purchase': 'DeepOrange',
        'proforma': 'Purple',
        'order_purchase': 'Teal',
        'return_sale': 'Red',
        'return_purchase': 'Teal',
        'stock_in': 'Green',
        'stock_out': 'Brown',
        'transfer': 'Purple',
        'initial_stock': 'Teal',
        'manage_products': 'Blue',
        'client_payment': 'Teal',
        'supplier_payment': 'Brown',
    }

    PAY_DEFAULT = ''
    PAY_CASH = 'Espèces'
    PAY_CHECK = 'Chèque'
    PAY_WIRE = 'Virement'
    PAY_DEPOSIT = 'Versement'

    PAYMENT_METHODS = (
        {'label': 'Par défaut', 'value': PAY_DEFAULT},
        {'label': 'Espèces', 'value': PAY_CASH},
        {'label': 'Chèque', 'value': PAY_CHECK},
        {'label': 'Virement', 'value': PAY_WIRE},
        {'label': 'Versement', 'value': PAY_DEPOSIT},
    )

    EXPENSE_CATEGORIES = (
        'Autre',
        'Assurance',
        'Banque / Frais bancaires',
        'Comptabilité / Juridique',
        'Électricité / Gaz / Eau',
        'Emballage / Sachets',
        'Équipement / Matériel',
        'Fournitures',
        'Frais divers',
        'Impôts / Taxes',
        'Internet / Téléphone',
        'Livraison / Expédition',
        'Loyer',
        'Maintenance',
        'Nettoyage / Hygiène',
        'Publicité / Marketing',
        'Restauration / Café',
        'Salaires / Avances',
        'Sécurité / Gardiennage',
        'Transport / Carburant',
        'Voyage / Déplacement',
    )

    EXPENSE_CATEGORY_STYLES = {
        'Autre': {'icon': 'cash-minus', 'color': (0.80, 0.14, 0.14, 1), 'bg': (1, 0.93, 0.93, 1)},
        'Assurance': {'icon': 'shield-check-outline', 'color': (0.08, 0.45, 0.82, 1), 'bg': (0.90, 0.95, 1, 1)},
        'Banque / Frais bancaires': {'icon': 'bank-outline', 'color': (0.18, 0.35, 0.68, 1), 'bg': (0.91, 0.94, 1, 1)},
        'Comptabilité / Juridique': {'icon': 'scale-balance', 'color': (0.42, 0.25, 0.68, 1), 'bg': (0.95, 0.92, 1, 1)},
        'Électricité / Gaz / Eau': {'icon': 'lightning-bolt-outline', 'color': (0.88, 0.52, 0.0, 1), 'bg': (1, 0.96, 0.88, 1)},
        'Emballage / Sachets': {'icon': 'package-variant', 'color': (0.62, 0.38, 0.12, 1), 'bg': (0.98, 0.94, 0.89, 1)},
        'Équipement / Matériel': {'icon': 'tools', 'color': (0.28, 0.38, 0.48, 1), 'bg': (0.92, 0.94, 0.97, 1)},
        'Fournitures': {'icon': 'package-variant-closed', 'color': (0.0, 0.55, 0.55, 1), 'bg': (0.90, 0.98, 0.98, 1)},
        'Frais divers': {'icon': 'dots-horizontal-circle-outline', 'color': (0.42, 0.46, 0.54, 1), 'bg': (0.94, 0.95, 0.97, 1)},
        'Impôts / Taxes': {'icon': 'receipt-text-outline', 'color': (0.76, 0.16, 0.26, 1), 'bg': (1, 0.92, 0.93, 1)},
        'Internet / Téléphone': {'icon': 'wifi', 'color': (0.0, 0.48, 0.78, 1), 'bg': (0.90, 0.96, 1, 1)},
        'Livraison / Expédition': {'icon': 'truck-delivery', 'color': (0.85, 0.45, 0.0, 1), 'bg': (1, 0.95, 0.89, 1)},
        'Loyer': {'icon': 'home-city-outline', 'color': (0.35, 0.28, 0.78, 1), 'bg': (0.93, 0.92, 1, 1)},
        'Maintenance': {'icon': 'wrench-outline', 'color': (0.55, 0.32, 0.12, 1), 'bg': (0.98, 0.94, 0.90, 1)},
        'Nettoyage / Hygiène': {'icon': 'broom', 'color': (0.04, 0.58, 0.48, 1), 'bg': (0.90, 0.98, 0.96, 1)},
        'Publicité / Marketing': {'icon': 'bullhorn-outline', 'color': (0.68, 0.15, 0.58, 1), 'bg': (0.99, 0.92, 0.98, 1)},
        'Restauration / Café': {'icon': 'coffee-outline', 'color': (0.52, 0.30, 0.15, 1), 'bg': (0.97, 0.94, 0.91, 1)},
        'Salaires / Avances': {'icon': 'account-cash-outline', 'color': (0.04, 0.58, 0.25, 1), 'bg': (0.90, 0.98, 0.92, 1)},
        'Sécurité / Gardiennage': {'icon': 'shield-lock-outline', 'color': (0.20, 0.30, 0.55, 1), 'bg': (0.91, 0.93, 0.98, 1)},
        'Transport / Carburant': {'icon': 'gas-station-outline', 'color': (0.08, 0.45, 0.82, 1), 'bg': (0.90, 0.95, 1, 1)},
        'Voyage / Déplacement': {'icon': 'car', 'color': (0.0, 0.52, 0.62, 1), 'bg': (0.90, 0.97, 0.99, 1)},
    }

    @classmethod
    def get_expense_style(cls, category: str) -> dict:
        return cls.EXPENSE_CATEGORY_STYLES.get(
            str(category or '').strip(),
            cls.EXPENSE_CATEGORY_STYLES['Autre'],
        )

    @classmethod
    def get_doc_type_for_mode(cls, mode: str) -> str:
        return cls.MODE_TO_DOC.get(str(mode or '').lower(), cls.DOC_BV)

    @classmethod
    def get_entity_type(cls, mode_or_doc: str) -> str:
        val = str(mode_or_doc or '').upper()
        if val in cls.PURCHASE_DOC_TYPES or val == cls.DOC_SUPPLIER_PAY:
            return cls.ENTITY_SUPPLIER
        doc = cls.MODE_TO_DOC.get(str(mode_or_doc or '').lower(), '')
        if doc in cls.PURCHASE_DOC_TYPES or doc == cls.DOC_SUPPLIER_PAY:
            return cls.ENTITY_SUPPLIER
        return cls.ENTITY_CLIENT

    @classmethod
    def normalize_price_category(cls, cat: str) -> str:
        clean = str(cat or '').strip()
        if clean in cls.PRICE_CATEGORIES:
            return clean
        return cls.PRICE_CAT_DETAIL

    @classmethod
    def is_default_entity(cls, entity_id=None, entity_type: str = '') -> bool:
        try:
            eid = int(entity_id) if entity_id is not None else None
        except (ValueError, TypeError):
            eid = None
        if entity_type == cls.ENTITY_CLIENT:
            return eid == cls.DEFAULT_CLIENT_ID
        if entity_type == cls.ENTITY_SUPPLIER:
            return eid == cls.DEFAULT_SUPPLIER_ID
        return eid in (cls.DEFAULT_CLIENT_ID, cls.DEFAULT_SUPPLIER_ID)

    @classmethod
    def is_default_product(cls, item_or_id) -> bool:
        if isinstance(item_or_id, dict):
            pid = item_or_id.get('id', item_or_id.get('product_id'))
        else:
            pid = item_or_id
        try:
            return int(pid) == cls.DEFAULT_PRODUCT_ID
        except (ValueError, TypeError):
            return False

    @classmethod
    def is_unlimited_stock(cls, qty) -> bool:
        return FinancialEngine.to_decimal(qty) <= Decimal(str(cls.UNLIMITED_STOCK_THRESHOLD))


# ==============================================================================
# 3. MOTEUR DE CALCUL UNIQUE (100% PARITÉ MAGPRO DESKTOP)
# ==============================================================================
class FinancialEngine:
    """Fonction de calcul unique pour panier, paiement, stats, PDF et tickets."""

    TWO_PLACES = Decimal('0.01')
    FOUR_PLACES = Decimal('0.0001')

    @staticmethod
    def to_decimal(value, default: str = '0.00') -> Decimal:
        if value is None or value == '':
            return Decimal(default)
        if isinstance(value, Decimal):
            return value if value.is_finite() else Decimal(default)
        try:
            clean = str(value).strip().replace(' ', '').replace(',', '.')
            d = Decimal(clean)
            return d if d.is_finite() else Decimal(default)
        except (InvalidOperation, ValueError, TypeError):
            return Decimal(default)

    @classmethod
    def quantize_money(cls, value) -> Decimal:
        return cls.to_decimal(value).quantize(cls.TWO_PLACES, rounding=ROUND_HALF_UP)

    @classmethod
    def quantize_qty(cls, value) -> Decimal:
        return cls.to_decimal(value).quantize(cls.FOUR_PLACES, rounding=ROUND_HALF_UP)

    @classmethod
    def format_money(cls, value, with_currency: bool = True) -> str:
        d = cls.quantize_money(value)
        formatted = f'{d:,.2f}'.replace(',', ' ').replace('.', ',')
        return f'{formatted} {AppConstants.CURRENCY}' if with_currency else formatted

    @classmethod
    def format_qty(cls, value) -> str:
        d = cls.to_decimal(value)
        if d == d.to_integral_value():
            return str(int(d))
        return f'{d.normalize():f}'

    @classmethod
    def validate_item_bounds(cls, qty, price=0, tva=0) -> tuple:
        q = cls.to_decimal(qty)
        p = cls.to_decimal(price)
        t = cls.to_decimal(tva)
        if q <= Decimal('0') or q > AppConstants.MAX_QTY_LIMIT:
            return False, f'Quantité invalide : doit être comprise entre 0 et {int(AppConstants.MAX_QTY_LIMIT):,}.'
        if p < Decimal('0') or p >= AppConstants.MAX_AMOUNT_LIMIT:
            return False, 'Prix unitaire invalide ou hors limites autorisées.'
        if t < Decimal('0') or t > AppConstants.MAX_PERCENT_LIMIT:
            return False, 'Le taux de TVA doit être compris entre 0 % et 100 %.'
        return True, ''

    @classmethod
    def validate_transaction_bounds(cls, total_amount, paid_amount=0) -> tuple:
        tot = cls.to_decimal(total_amount)
        paid = cls.to_decimal(paid_amount)
        if abs(tot) >= AppConstants.MAX_AMOUNT_LIMIT:
            return False, 'Le montant total dépasse la limite autorisée (1 000 000 000 DA).'
        if abs(paid) >= AppConstants.MAX_AMOUNT_LIMIT:
            return False, 'Le montant payé dépasse la limite autorisée.'
        return True, ''

    @classmethod
    def calculate_document_totals(
        cls,
        items,
        doc_type: str = AppConstants.DOC_BV,
        global_discount=Decimal('0.00'),
        payment_method: str = AppConstants.PAY_DEFAULT,
        is_comptoir: bool = False,
    ) -> dict:
        if doc_type in AppConstants.QTY_ONLY_DOC_TYPES:
            items_iter = items.values() if isinstance(items, dict) else (items or [])
            real_cnt = sum(1 for it in items_iter if not AppConstants.is_default_product(it))
            z = Decimal('0.00')
            return {
                'item_count_real': real_cnt,
                'total_gross_ht': z,
                'total_ht_net_of_lines': z,
                'total_tva_brut_raw': z,
                'total_item_discounts': z,
                'global_discount': z,
                'total_global_discount': z,
                'total_tva_final': z,
                'total_ht_final': z,
                'total_all_discounts_final': z,
                'total_discounts_ht': z,
                'net_to_pay': z,
                'stamp_duty': z,
                'show_stamp_duty_widget': False,
                'final_total': z,
            }

        include_tva = doc_type in AppConstants.INVOICE_TAX_DOC_TYPES
        include_stamp = doc_type in AppConstants.STAMP_DUTY_DOC_TYPES

        total_gross_ht = Decimal('0.00')
        total_ht_net_of_lines = Decimal('0.00')
        total_tva_brut_raw = Decimal('0.00')
        total_item_discounts = Decimal('0.00')
        item_count_real = 0

        items_iterable = items.values() if isinstance(items, dict) else (items or [])
        for item in items_iterable:
            if not AppConstants.is_default_product(item):
                item_count_real += 1
            price = max(Decimal('0.00'), cls.to_decimal(item.get('price', 0)))
            qty = cls.to_decimal(item.get('qty', item.get('quantity', 0)))
            tva_percent = max(Decimal('0.00'), cls.to_decimal(item.get('tva', 0))) if include_tva else Decimal('0.00')
            disc_amt = max(Decimal('0.00'), cls.to_decimal(item.get('discount_amount', 0)))
            discount_type = item.get('discount_type', 'fixed')

            line_gross = cls.quantize_money(price * qty)
            total_gross_ht += line_gross

            unit_discount = disc_amt if discount_type == 'fixed' else price * (disc_amt / Decimal('100'))
            price_after_disc = price - unit_discount
            line_net = cls.quantize_money(price_after_disc * qty)
            total_item_discounts += abs(line_gross - line_net)
            total_ht_net_of_lines += line_net

            line_tva = cls.quantize_money(line_net * (tva_percent / Decimal('100')))
            total_tva_brut_raw += line_tva

        total_global_discount = max(Decimal('0.00'), cls.to_decimal(global_discount))
        total_ttc_before_discount = total_ht_net_of_lines + total_tva_brut_raw

        if total_ttc_before_discount < Decimal('0'):
            if total_global_discount > abs(total_ttc_before_discount):
                total_global_discount = abs(total_ttc_before_discount)
            net_to_pay = total_ttc_before_discount + total_global_discount
        else:
            if total_global_discount > total_ttc_before_discount:
                total_global_discount = total_ttc_before_discount
            net_to_pay = total_ttc_before_discount - total_global_discount

        if abs(total_ttc_before_discount) > Decimal('0'):
            ratio = net_to_pay / total_ttc_before_discount
        else:
            ratio = Decimal('1')

        total_tva_final = cls.quantize_money(total_tva_brut_raw * ratio)
        total_ht_final = cls.quantize_money(net_to_pay - total_tva_final)
        total_all_discounts_final = cls.quantize_money(total_item_discounts + total_global_discount)
        global_discount_ht = abs(total_ht_net_of_lines - total_ht_final)
        total_discounts_ht = cls.quantize_money(total_item_discounts + global_discount_ht)

        stamp_duty = Decimal('0.00')
        show_stamp_duty_widget = False
        if not is_comptoir and include_stamp and payment_method == AppConstants.PAY_CASH:
            show_stamp_duty_widget = True
            if net_to_pay > Decimal('300'):
                if net_to_pay <= Decimal('500000'):
                    duty = net_to_pay * Decimal('0.01')
                elif net_to_pay <= Decimal('1000000'):
                    duty = net_to_pay * Decimal('0.015')
                else:
                    duty = net_to_pay * Decimal('0.02')
                stamp_duty = max(Decimal('5.00'), cls.quantize_money(duty))

        final_total = cls.quantize_money(net_to_pay + stamp_duty)
        return {
            'item_count_real': item_count_real,
            'total_gross_ht': cls.quantize_money(total_gross_ht),
            'total_ht_net_of_lines': cls.quantize_money(total_ht_net_of_lines),
            'total_tva_brut_raw': cls.quantize_money(total_tva_brut_raw),
            'total_item_discounts': cls.quantize_money(total_item_discounts),
            'global_discount': cls.quantize_money(total_global_discount),
            'total_global_discount': cls.quantize_money(total_global_discount),
            'total_tva_final': total_tva_final,
            'total_ht_final': total_ht_final,
            'total_all_discounts_final': total_all_discounts_final,
            'total_discounts_ht': total_discounts_ht,
            'net_to_pay': cls.quantize_money(net_to_pay),
            'stamp_duty': stamp_duty,
            'show_stamp_duty_widget': show_stamp_duty_widget,
            'final_total': final_total,
        }

    @classmethod
    def number_to_words_fr(cls, n_in) -> str:
        n_decimal = cls.quantize_money(n_in)
        if n_decimal.is_zero():
            return 'zéro'
        is_negative = n_decimal < 0
        if is_negative:
            n_decimal = abs(n_decimal)
        integer_part = int(n_decimal)
        decimal_part = int((n_decimal - integer_part) * 100)

        units = ['', 'un', 'deux', 'trois', 'quatre', 'cinq', 'six', 'sept', 'huit', 'neuf']
        teens = ['dix', 'onze', 'douze', 'treize', 'quatorze', 'quinze', 'seize', 'dix-sept', 'dix-huit', 'dix-neuf']
        tens = ['', 'dix', 'vingt', 'trente', 'quarante', 'cinquante', 'soixante', 'soixante-dix', 'quatre-vingt', 'quatre-vingt-dix']

        def get_words(num: int) -> list:
            if num == 0:
                return []
            if num < 10:
                return [units[num]]
            if num < 20:
                return [teens[num - 10]]
            if num < 70:
                word_parts = [tens[num // 10]]
                rem = num % 10
                if rem > 0:
                    word_parts.append('-et-un' if (rem == 1 and num // 10 not in (7, 9)) else f'-{units[rem]}')
                return [''.join(word_parts)]
            if num < 80:
                return [f'soixante-{get_words(num - 60)[0]}']
            if num < 100:
                rem = num - 80
                return ['quatre-vingts'] if rem == 0 else [f'quatre-vingt-{get_words(rem)[0]}']
            if num < 1000:
                p_val = num // 100
                prefix = [] if p_val == 1 else get_words(p_val)
                rest = get_words(num % 100)
                base = 'cents' if (p_val > 1 and not rest) else 'cent'
                return prefix + [base] + rest
            if num < 1000000:
                p_val = num // 1000
                prefix = [] if p_val == 1 else get_words(p_val)
                return prefix + ['mille'] + get_words(num % 1000)
            if num < 2000000:
                return ['un', 'million'] + get_words(num % 1000000)
            if num < 1000000000:
                return get_words(num // 1000000) + ['millions'] + get_words(num % 1000000)
            return ['nombre', 'trop', 'grand']

        int_words = get_words(integer_part) if integer_part > 0 else []
        dec_words = get_words(decimal_part) if decimal_part > 0 else []
        res = []
        if is_negative:
            res.append('moins')
        if int_words:
            res.extend(int_words)
        if dec_words:
            if int_words:
                res.append('et')
            res.extend(dec_words)
            res.append('centimes' if decimal_part > 1 else 'centime')
        return ' '.join(res).strip().replace(' -', '-')


# ==============================================================================
# 4. BASE DE DONNÉES SQLITE AUTONOME (SCHÉMA PROPRE & RECALCUL CHRONOLOGIQUE)
# ==============================================================================
class DatabaseManager:
    def __init__(self, db_name: str = AppConstants.DB_NAME):
        if platform == 'android' and PythonActivity is not None:
            try:
                files_dir = PythonActivity.mActivity.getFilesDir().getAbsolutePath()
                self.db_path = os.path.join(files_dir, db_name)
            except Exception:
                self.db_path = os.path.join(APP_DIR, db_name)
        else:
            self.db_path = os.path.join(APP_DIR, db_name)

        self.cleanup_wal_files()
        self.create_tables()
        self.seed_data()

    def get_connection(self) -> sqlite3.Connection:
        db_dir = os.path.dirname(self.db_path)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir, exist_ok=True)
        conn = sqlite3.connect(self.db_path, timeout=15)
        conn.row_factory = sqlite3.Row
        try:
            conn.execute('PRAGMA journal_mode=WAL;')
            conn.execute('PRAGMA synchronous=NORMAL;')
            conn.execute('PRAGMA foreign_keys=ON;')
        except Exception:
            pass
        return conn

    def cleanup_wal_files(self):
        wal = self.db_path + '-wal'
        shm = self.db_path + '-shm'
        if os.path.exists(wal) or os.path.exists(shm):
            try:
                conn = sqlite3.connect(self.db_path)
                conn.execute('PRAGMA wal_checkpoint(TRUNCATE);')
                conn.execute('PRAGMA journal_mode=DELETE;')
                conn.commit()
                conn.close()
            except Exception:
                pass
            for f in (wal, shm):
                if os.path.exists(f):
                    try:
                        os.remove(f)
                    except Exception:
                        pass

    def create_tables(self):
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            cur.execute('''
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    role TEXT NOT NULL
                )
            ''')
            cur.execute('''
                CREATE TABLE IF NOT EXISTS product_families (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL
                )
            ''')
            cur.execute('''
                CREATE TABLE IF NOT EXISTS products (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL,
                    barcode TEXT DEFAULT '',
                    product_ref TEXT DEFAULT '',
                    reference TEXT DEFAULT '',
                    family TEXT DEFAULT '',
                    purchase_price REAL DEFAULT 0,
                    price REAL DEFAULT 0,
                    price_semi REAL DEFAULT 0,
                    price_wholesale REAL DEFAULT 0,
                    tva REAL DEFAULT 0,
                    stock_store REAL DEFAULT 0,
                    stock_warehouse REAL DEFAULT 0,
                    image_path TEXT DEFAULT '',
                    is_promo_active INTEGER DEFAULT 0,
                    promo_type TEXT DEFAULT 'fixed',
                    promo_value REAL DEFAULT 0,
                    promo_expiry TEXT DEFAULT ''
                )
            ''')
            cur.execute('''
                CREATE TABLE IF NOT EXISTS entities (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    entity_type TEXT NOT NULL,
                    name TEXT NOT NULL,
                    phone TEXT DEFAULT '',
                    balance REAL DEFAULT 0,
                    address TEXT DEFAULT '',
                    email TEXT DEFAULT '',
                    activity TEXT DEFAULT '',
                    rc TEXT DEFAULT '',
                    nif TEXT DEFAULT '',
                    nis TEXT DEFAULT '',
                    nai TEXT DEFAULT '',
                    rib TEXT DEFAULT '',
                    price_category TEXT DEFAULT 'Détail',
                    gps_location TEXT DEFAULT '',
                    UNIQUE(name, entity_type)
                )
            ''')
            cur.execute('''
                CREATE TABLE IF NOT EXISTS transactions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    doc_type TEXT NOT NULL,
                    entity_category TEXT NOT NULL,
                    entity_id INTEGER,
                    doc_number TEXT DEFAULT '',
                    total_amount REAL DEFAULT 0,
                    global_discount REAL DEFAULT 0,
                    paid_amount REAL DEFAULT 0,
                    payment_method TEXT DEFAULT '',
                    stamp_duty REAL DEFAULT 0,
                    location TEXT DEFAULT 'store',
                    note TEXT DEFAULT '',
                    user_name TEXT DEFAULT 'ADMIN',
                    date TIMESTAMP NOT NULL,
                    FOREIGN KEY(entity_id) REFERENCES entities(id)
                )
            ''')
            cur.execute('''
                CREATE TABLE IF NOT EXISTS transaction_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    transaction_id INTEGER NOT NULL,
                    product_id INTEGER NOT NULL,
                    product_name TEXT NOT NULL,
                    qty REAL NOT NULL,
                    price REAL NOT NULL,
                    cost_at_time REAL DEFAULT 0,
                    tva REAL DEFAULT 0,
                    discount_amount REAL DEFAULT 0,
                    discount_type TEXT DEFAULT 'fixed',
                    FOREIGN KEY(transaction_id) REFERENCES transactions(id) ON DELETE CASCADE,
                    FOREIGN KEY(product_id) REFERENCES products(id)
                )
            ''')
            cur.execute('''
                CREATE TABLE IF NOT EXISTS stock_transfers (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    transfer_number TEXT UNIQUE NOT NULL,
                    transfer_date TIMESTAMP NOT NULL,
                    source_location TEXT NOT NULL,
                    destination_location TEXT NOT NULL,
                    notes TEXT DEFAULT '',
                    user_name TEXT DEFAULT 'ADMIN'
                )
            ''')
            cur.execute('''
                CREATE TABLE IF NOT EXISTS stock_transfer_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    transfer_id INTEGER NOT NULL,
                    product_id INTEGER NOT NULL,
                    product_name TEXT NOT NULL,
                    qty REAL NOT NULL,
                    FOREIGN KEY(transfer_id) REFERENCES stock_transfers(id) ON DELETE CASCADE,
                    FOREIGN KEY(product_id) REFERENCES products(id)
                )
            ''')
            cur.execute('''
                CREATE TABLE IF NOT EXISTS general_expenses (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    description TEXT DEFAULT '',
                    amount REAL NOT NULL,
                    category TEXT DEFAULT 'Autre',
                    expense_date TIMESTAMP NOT NULL,
                    recorded_by TEXT DEFAULT 'ADMIN'
                )
            ''')
            cur.execute('''
                CREATE TABLE IF NOT EXISTS app_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )
            ''')

            expected_columns = {
                'products': (
                    ("barcode", "TEXT DEFAULT ''"),
                    ("product_ref", "TEXT DEFAULT ''"),
                    ("reference", "TEXT DEFAULT ''"),
                    ("family", "TEXT DEFAULT ''"),
                    ("purchase_price", "REAL DEFAULT 0"),
                    ("price", "REAL DEFAULT 0"),
                    ("price_semi", "REAL DEFAULT 0"),
                    ("price_wholesale", "REAL DEFAULT 0"),
                    ("tva", "REAL DEFAULT 0"),
                    ("stock_store", "REAL DEFAULT 0"),
                    ("stock_warehouse", "REAL DEFAULT 0"),
                    ("image_path", "TEXT DEFAULT ''"),
                    ("is_promo_active", "INTEGER DEFAULT 0"),
                    ("promo_type", "TEXT DEFAULT 'fixed'"),
                    ("promo_value", "REAL DEFAULT 0"),
                    ("promo_expiry", "TEXT DEFAULT ''"),
                ),
                'entities': (
                    ("phone", "TEXT DEFAULT ''"),
                    ("balance", "REAL DEFAULT 0"),
                    ("address", "TEXT DEFAULT ''"),
                    ("email", "TEXT DEFAULT ''"),
                    ("activity", "TEXT DEFAULT ''"),
                    ("rc", "TEXT DEFAULT ''"),
                    ("nif", "TEXT DEFAULT ''"),
                    ("nis", "TEXT DEFAULT ''"),
                    ("nai", "TEXT DEFAULT ''"),
                    ("rib", "TEXT DEFAULT ''"),
                    ("price_category", "TEXT DEFAULT 'Détail'"),
                    ("gps_location", "TEXT DEFAULT ''"),
                ),
                'transactions': (
                    ("doc_number", "TEXT DEFAULT ''"),
                    ("total_amount", "REAL DEFAULT 0"),
                    ("global_discount", "REAL DEFAULT 0"),
                    ("paid_amount", "REAL DEFAULT 0"),
                    ("payment_method", "TEXT DEFAULT ''"),
                    ("stamp_duty", "REAL DEFAULT 0"),
                    ("location", "TEXT DEFAULT 'store'"),
                    ("note", "TEXT DEFAULT ''"),
                    ("user_name", "TEXT DEFAULT 'ADMIN'"),
                ),
                'transaction_items': (
                    ("cost_at_time", "REAL DEFAULT 0"),
                    ("tva", "REAL DEFAULT 0"),
                    ("discount_amount", "REAL DEFAULT 0"),
                    ("discount_type", "TEXT DEFAULT 'fixed'"),
                ),
                'general_expenses': (
                    ("description", "TEXT DEFAULT ''"),
                    ("category", "TEXT DEFAULT 'Autre'"),
                    ("recorded_by", "TEXT DEFAULT 'ADMIN'"),
                ),
            }
            for tbl_name, cols_spec in expected_columns.items():
                existing = {r[1] for r in cur.execute(f'PRAGMA table_info({tbl_name})').fetchall()}
                for col_name, col_def in cols_spec:
                    if col_name not in existing:
                        cur.execute(f'ALTER TABLE {tbl_name} ADD COLUMN {col_name} {col_def}')

            for idx_sql in (
                'CREATE INDEX IF NOT EXISTS idx_prod_name ON products(name)',
                'CREATE INDEX IF NOT EXISTS idx_prod_barcode ON products(barcode)',
                'CREATE INDEX IF NOT EXISTS idx_prod_family ON products(family)',
                'CREATE INDEX IF NOT EXISTS idx_ent_type ON entities(entity_type)',
                'CREATE INDEX IF NOT EXISTS idx_tx_date ON transactions(date)',
                'CREATE INDEX IF NOT EXISTS idx_tx_doc ON transactions(doc_type)',
                'CREATE INDEX IF NOT EXISTS idx_tx_ent ON transactions(entity_id)',
                'CREATE INDEX IF NOT EXISTS idx_txi_tx ON transaction_items(transaction_id)',
                'CREATE INDEX IF NOT EXISTS idx_txi_prod ON transaction_items(product_id)',
                'CREATE INDEX IF NOT EXISTS idx_exp_date ON general_expenses(expense_date)',
            ):
                cur.execute(idx_sql)
            conn.commit()
        finally:
            conn.close()

    def seed_data(self):
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            cur.execute('SELECT COUNT(*) FROM users')
            if cur.fetchone()[0] == 0:
                pw_hash = hashlib.sha256(''.encode('utf-8')).hexdigest()
                cur.execute(
                    'INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)',
                    (AppConstants.ADMIN_USERNAME, pw_hash, 'admin'),
                )
            cur.execute(
                'INSERT OR IGNORE INTO entities (id, entity_type, name, price_category, balance) VALUES (?, ?, ?, ?, 0)',
                (AppConstants.DEFAULT_CLIENT_ID, AppConstants.ENTITY_CLIENT, AppConstants.DEFAULT_CLIENT_NAME, AppConstants.PRICE_CAT_DETAIL),
            )
            cur.execute(
                'INSERT OR IGNORE INTO entities (id, entity_type, name, price_category, balance) VALUES (?, ?, ?, ?, 0)',
                (AppConstants.DEFAULT_SUPPLIER_ID, AppConstants.ENTITY_SUPPLIER, AppConstants.DEFAULT_SUPPLIER_NAME, AppConstants.PRICE_CAT_GROS),
            )
            cur.execute(
                'INSERT OR IGNORE INTO products (id, name, stock_store, stock_warehouse) VALUES (?, ?, ?, ?)',
                (
                    AppConstants.DEFAULT_PRODUCT_ID,
                    AppConstants.DEFAULT_PRODUCT_NAME,
                    AppConstants.UNLIMITED_STOCK_VALUE,
                    AppConstants.UNLIMITED_STOCK_VALUE,
                ),
            )
            defaults = {
                'allow_negative_stock': 'True',
                'warn_price_below_cost': 'True',
                'printer_auto': 'True',
                'show_balance_in_pdf': 'False',
            }
            for k, v in defaults.items():
                cur.execute('INSERT OR IGNORE INTO app_settings (key, value) VALUES (?, ?)', (k, v))
            conn.commit()
        finally:
            conn.close()

    # --- Authentification & Paramètres ---
    def login(self, password: str) -> dict | None:
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            pw_hash = hashlib.sha256((password or '').encode('utf-8')).hexdigest()
            cur.execute('SELECT * FROM users WHERE username = ? AND password_hash = ?', (AppConstants.ADMIN_USERNAME, pw_hash))
            row = cur.fetchone()
            return dict(row) if row else None
        finally:
            conn.close()

    def update_admin_password(self, new_password: str) -> bool:
        conn = self.get_connection()
        try:
            pw_hash = hashlib.sha256((new_password or '').encode('utf-8')).hexdigest()
            conn.execute('UPDATE users SET password_hash = ? WHERE username = ?', (pw_hash, AppConstants.ADMIN_USERNAME))
            conn.commit()
            return True
        except Exception:
            return False
        finally:
            conn.close()

    def set_setting(self, key: str, value):
        conn = self.get_connection()
        try:
            conn.execute('REPLACE INTO app_settings (key, value) VALUES (?, ?)', (str(key), str(value)))
            conn.commit()
        finally:
            conn.close()

    def get_setting(self, key: str, default=None):
        conn = self.get_connection()
        try:
            row = conn.execute('SELECT value FROM app_settings WHERE key = ?', (str(key),)).fetchone()
            return row['value'] if row else default
        finally:
            conn.close()

    def get_bool_setting(self, key: str, default: bool = False) -> bool:
        raw = self.get_setting(key, 'True' if default else 'False')
        return str(raw).strip().lower() in ('true', '1', 'yes', 'on')

    def get_store_info(self) -> dict:
        conn = self.get_connection()
        try:
            rows = conn.execute("SELECT key, value FROM app_settings WHERE key LIKE 'store_%'").fetchall()
            return {r['key']: r['value'] for r in rows}
        finally:
            conn.close()

    # --- Familles ---
    def get_families(self) -> list:
        conn = self.get_connection()
        try:
            rows = conn.execute('SELECT name FROM product_families ORDER BY name COLLATE NOCASE ASC').fetchall()
            return [r['name'] for r in rows]
        finally:
            conn.close()

    def add_family(self, name: str) -> bool:
        clean = str(name or '').strip()
        if not clean or clean.upper() == AppConstants.FAM_ALL:
            return False
        conn = self.get_connection()
        try:
            conn.execute('INSERT INTO product_families (name) VALUES (?)', (clean,))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()

    def delete_family(self, name: str) -> bool:
        clean = str(name or '').strip()
        if not clean:
            return False
        conn = self.get_connection()
        try:
            if conn.execute('SELECT 1 FROM products WHERE family = ? LIMIT 1', (clean,)).fetchone():
                return False
            conn.execute('DELETE FROM product_families WHERE name = ?', (clean,))
            conn.commit()
            return True
        finally:
            conn.close()

    # --- Numérotation des documents ---
    def peek_next_doc_number(self, doc_type: str, cur=None) -> str:
        now = datetime.now()
        suffix = now.strftime('%Y') if doc_type in AppConstants.YEARLY_RESET_SEQUENCES else now.strftime('%d%m')
        close_after = False
        if cur is None:
            conn = self.get_connection()
            cur = conn.cursor()
            close_after = True
        else:
            conn = None
        try:
            if doc_type == AppConstants.DOC_TR:
                cur.execute(f"SELECT transfer_number FROM stock_transfers WHERE transfer_number LIKE 'TR%/{suffix}'")
            else:
                cur.execute(f"SELECT doc_number FROM transactions WHERE doc_number LIKE '{doc_type}%/{suffix}'")
            max_val = 0
            for r in cur.fetchall():
                val_str = str(r[0] or '').upper()
                if val_str.startswith(doc_type) and '/' in val_str:
                    num_part = val_str[len(doc_type):].split('/')[0]
                    if num_part.isdigit():
                        max_val = max(max_val, int(num_part))
            return f'{doc_type}{max_val + 1:05d}/{suffix}'
        finally:
            if close_after and conn:
                conn.close()

    def get_next_product_ref(self) -> str:
        conn = self.get_connection()
        try:
            max_num = 0
            for r in conn.execute('SELECT product_ref FROM products').fetchall():
                s = str(r[0] or '').strip()
                if s.isdigit():
                    max_num = max(max_num, int(s))
            return f'{max_num + 1:05d}'
        finally:
            conn.close()

    # --- Produits ---
    def get_products(self, limit: int = 50, offset: int = 0, search_query: str = '', family_filter: str = AppConstants.FAM_ALL, include_default: bool = False) -> list:
        conn = self.get_connection()
        try:
            conditions = []
            params = []
            if not include_default:
                conditions.append('id <> ?')
                params.append(AppConstants.DEFAULT_PRODUCT_ID)
            if search_query:
                q = str(search_query).strip()
                if q.lower() == 'promo':
                    conditions.append('is_promo_active = 1')
                else:
                    for term in q.split():
                        conditions.append('(name LIKE ? OR barcode LIKE ? OR product_ref LIKE ? OR reference LIKE ?)')
                        like = f'%{term}%'
                        params.extend([like, like, like, like])
            if family_filter and family_filter != AppConstants.FAM_ALL:
                conditions.append('family = ?')
                params.append(family_filter)

            where_sql = f"WHERE {' AND '.join(conditions)}" if conditions else ''
            sql = f'''
                SELECT * FROM products
                {where_sql}
                ORDER BY CASE WHEN id = {AppConstants.DEFAULT_PRODUCT_ID} THEN 0 ELSE 1 END, name COLLATE NOCASE ASC
                LIMIT ? OFFSET ?
            '''
            params.extend([limit, offset])
            return [dict(r) for r in conn.execute(sql, params).fetchall()]
        finally:
            conn.close()

    def get_product_by_id(self, p_id: int) -> dict | None:
        conn = self.get_connection()
        try:
            row = conn.execute('SELECT * FROM products WHERE id = ?', (p_id,)).fetchone()
            return dict(row) if row else None
        finally:
            conn.close()

    def get_product_by_barcode(self, code: str) -> dict | None:
        clean = str(code or '').strip()
        if not clean:
            return None
        conn = self.get_connection()
        try:
            row = conn.execute(
                'SELECT * FROM products WHERE barcode = ? OR product_ref = ? LIMIT 1',
                (clean, clean),
            ).fetchone()
            return dict(row) if row else None
        finally:
            conn.close()

    def get_last_product_price_for_entity(self, product_id: int, entity_id: int | None, doc_type_or_mode: str) -> Decimal | None:
        if not product_id or AppConstants.is_default_product(product_id) or not entity_id:
            return None
        val = str(doc_type_or_mode or '').upper()
        if val in (AppConstants.DOC_RF, 'RETURN_PURCHASE'):
            target_types = (AppConstants.DOC_BA, AppConstants.DOC_FF, AppConstants.DOC_BI)
            ent_cat = AppConstants.ENTITY_SUPPLIER
        elif val in (AppConstants.DOC_RC, 'RETURN_SALE'):
            target_types = (AppConstants.DOC_BV, AppConstants.DOC_FC)
            ent_cat = AppConstants.ENTITY_CLIENT
        else:
            return None

        if AppConstants.is_default_entity(entity_id, ent_cat):
            return None

        conn = self.get_connection()
        try:
            placeholders = ','.join('?' for _ in target_types)
            row = conn.execute(
                f'''
                SELECT ti.price, ti.discount_amount, ti.discount_type
                FROM transaction_items ti
                JOIN transactions t ON ti.transaction_id = t.id
                WHERE ti.product_id = ?
                  AND t.entity_id = ?
                  AND t.entity_category = ?
                  AND t.doc_type IN ({placeholders})
                ORDER BY t.date DESC, t.id DESC
                LIMIT 1
                ''',
                (int(product_id), int(entity_id), ent_cat, *target_types),
            ).fetchone()
            if not row:
                return None
            price = FinancialEngine.to_decimal(row['price'])
            d_amt = FinancialEngine.to_decimal(row['discount_amount'])
            d_type = row['discount_type'] or 'fixed'
            u_disc = d_amt if d_type == 'fixed' else price * (d_amt / Decimal('100'))
            return max(Decimal('0.00'), FinancialEngine.quantize_money(price - u_disc))
        finally:
            conn.close()

    def is_product_in_real_use(self, product_id: int) -> bool:
        if int(product_id) == AppConstants.DEFAULT_PRODUCT_ID:
            return True
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            if cur.execute('SELECT 1 FROM stock_transfer_items WHERE product_id = ? LIMIT 1', (product_id,)).fetchone():
                return True
            row = cur.execute(
                '''
                SELECT 1 FROM transaction_items ti
                JOIN transactions t ON ti.transaction_id = t.id
                WHERE ti.product_id = ?
                  AND t.doc_type NOT IN (?, ?, ?)
                LIMIT 1
                ''',
                (product_id, AppConstants.DOC_FP, AppConstants.DOC_DP, AppConstants.DOC_BI),
            ).fetchone()
            return row is not None
        finally:
            conn.close()

    def get_product_bi_transaction(self, product_id: int) -> int | None:
        conn = self.get_connection()
        try:
            row = conn.execute(
                '''
                SELECT t.id FROM transaction_items ti
                JOIN transactions t ON ti.transaction_id = t.id
                WHERE ti.product_id = ? AND t.doc_type = ?
                LIMIT 1
                ''',
                (product_id, AppConstants.DOC_BI),
            ).fetchone()
            return row['id'] if row else None
        finally:
            conn.close()

    def save_product(self, product_data: dict, user_name: str = AppConstants.ADMIN_USERNAME) -> int:
        action = product_data.get('action', 'add')
        p_id = product_data.get('id')
        name = str(product_data.get('name', '')).strip()
        if not name:
            raise ValueError('La désignation est obligatoire.')
        if name.lower() == AppConstants.DEFAULT_PRODUCT_NAME.lower() or (p_id is not None and int(p_id) == AppConstants.DEFAULT_PRODUCT_ID):
            raise ValueError('Action interdite sur le produit par défaut.')

        barcode = str(product_data.get('barcode', '')).strip()
        product_ref = str(product_data.get('product_ref', '')).strip()
        reference = str(product_data.get('reference', '')).strip()
        family = str(product_data.get('family', '')).strip()
        if family == AppConstants.FAM_ALL:
            family = ''

        cost = FinancialEngine.quantize_money(product_data.get('purchase_price', 0))
        p1 = FinancialEngine.quantize_money(product_data.get('price', 0))
        p2 = FinancialEngine.quantize_money(product_data.get('price_semi', 0))
        p3 = FinancialEngine.quantize_money(product_data.get('price_wholesale', 0))
        tva = FinancialEngine.quantize_money(product_data.get('tva', 0))
        if any(v < 0 or v >= AppConstants.MAX_AMOUNT_LIMIT for v in (cost, p1, p2, p3)) or tva < 0 or tva > AppConstants.MAX_PERCENT_LIMIT:
            raise ValueError('Prix ou TVA hors limites autorisées.')

        is_unlimited = bool(product_data.get('is_unlimited', False))
        raw_stock = FinancialEngine.quantize_qty(product_data.get('stock_store', 0))
        if not is_unlimited and (raw_stock < 0 or raw_stock > AppConstants.MAX_QTY_LIMIT):
            raise ValueError('Quantité de stock initiale invalide.')
        if not is_unlimited and raw_stock > 0 and (raw_stock * cost) >= AppConstants.MAX_AMOUNT_LIMIT:
            raise ValueError('La valeur initiale du stock dépasse la limite autorisée.')

        is_promo = 1 if product_data.get('is_promo_active', False) else 0
        promo_type = 'percent' if product_data.get('promo_type') == 'percent' else 'fixed'
        promo_val = FinancialEngine.quantize_money(product_data.get('promo_value', 0))
        if is_promo:
            if promo_val < 0 or (promo_type == 'percent' and promo_val > AppConstants.MAX_PERCENT_LIMIT) or (promo_type == 'fixed' and promo_val >= AppConstants.MAX_AMOUNT_LIMIT):
                raise ValueError('Valeur promotionnelle hors limites.')
        promo_exp = str(product_data.get('promo_expiry', '')).strip()
        image_path = str(product_data.get('image_path', '')).strip()

        has_movements = self.is_product_in_real_use(p_id) if (action != 'add' and p_id is not None) else False

        conn = self.get_connection()
        try:
            cur = conn.cursor()
            cur.execute('BEGIN')
            if barcode:
                dup_bc = cur.execute(
                    'SELECT id, name FROM products WHERE barcode = ? AND id <> ?',
                    (barcode, p_id or -1),
                ).fetchone()
                if dup_bc:
                    raise ValueError(f"Ce code-barres est déjà utilisé par '{dup_bc['name']}'.")

            if family:
                cur.execute('INSERT OR IGNORE INTO product_families (name) VALUES (?)', (family,))

            if action == 'add':
                init_store = float(AppConstants.UNLIMITED_STOCK_VALUE) if is_unlimited else float(raw_stock)
                init_wh = float(AppConstants.UNLIMITED_STOCK_VALUE) if is_unlimited else 0.0
                cur.execute(
                    '''
                    INSERT INTO products (
                        name, barcode, product_ref, reference, family,
                        purchase_price, price, price_semi, price_wholesale, tva,
                        stock_store, stock_warehouse, image_path,
                        is_promo_active, promo_type, promo_value, promo_expiry
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''',
                    (
                        name, barcode, product_ref, reference, family,
                        float(cost), float(p1), float(p2), float(p3), float(tva),
                        init_store, init_wh, image_path,
                        is_promo, promo_type, float(promo_val), promo_exp,
                    ),
                )
                p_id = cur.lastrowid
                has_movements = False
            elif has_movements:
                cur.execute(
                    '''
                    UPDATE products SET
                        name=?, barcode=?, product_ref=?, reference=?, family=?,
                        price=?, price_semi=?, price_wholesale=?, tva=?, image_path=?,
                        is_promo_active=?, promo_type=?, promo_value=?, promo_expiry=?
                    WHERE id=?
                    ''',
                    (
                        name, barcode, product_ref, reference, family,
                        float(p1), float(p2), float(p3), float(tva), image_path,
                        is_promo, promo_type, float(promo_val), promo_exp, p_id,
                    ),
                )
            else:
                upd_store = float(AppConstants.UNLIMITED_STOCK_VALUE) if is_unlimited else float(raw_stock)
                upd_wh = float(AppConstants.UNLIMITED_STOCK_VALUE) if is_unlimited else 0.0
                cur.execute(
                    '''
                    UPDATE products SET
                        name=?, barcode=?, product_ref=?, reference=?, family=?,
                        purchase_price=?, price=?, price_semi=?, price_wholesale=?, tva=?,
                        stock_store=?, stock_warehouse=?, image_path=?,
                        is_promo_active=?, promo_type=?, promo_value=?, promo_expiry=?
                    WHERE id=?
                    ''',
                    (
                        name, barcode, product_ref, reference, family,
                        float(cost), float(p1), float(p2), float(p3), float(tva),
                        upd_store, upd_wh, image_path,
                        is_promo, promo_type, float(promo_val), promo_exp, p_id,
                    ),
                )

            if not has_movements:
                bi_row = cur.execute(
                    '''
                    SELECT t.id FROM transaction_items ti
                    JOIN transactions t ON ti.transaction_id = t.id
                    WHERE ti.product_id = ? AND t.doc_type = ?
                    LIMIT 1
                    ''',
                    (p_id, AppConstants.DOC_BI),
                ).fetchone()
                bi_id = bi_row['id'] if bi_row else None

                if is_unlimited or raw_stock == 0:
                    if bi_id:
                        self._delete_invoice_in_cursor(cur, bi_id)
                else:
                    total_bi = FinancialEngine.quantize_money(raw_stock * cost)
                    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    if bi_id:
                        cur.execute(
                            'UPDATE transactions SET total_amount = ?, paid_amount = ? WHERE id = ?',
                            (float(total_bi), float(total_bi), bi_id),
                        )
                        cur.execute(
                            '''
                            UPDATE transaction_items
                            SET product_name = ?, qty = ?, price = ?, cost_at_time = ?
                            WHERE transaction_id = ? AND product_id = ?
                            ''',
                            (name, float(raw_stock), float(cost), float(cost), bi_id, p_id),
                        )
                    else:
                        bi_num = self.peek_next_doc_number(AppConstants.DOC_BI, cur)
                        cur.execute(
                            '''
                            INSERT INTO transactions (
                                doc_type, entity_category, entity_id, doc_number,
                                total_amount, global_discount, paid_amount, payment_method,
                                stamp_duty, location, note, user_name, date
                            ) VALUES (?, ?, ?, ?, ?, 0, ?, 'Initial', 0, ?, 'Stock Initial', ?, ?)
                            ''',
                            (
                                AppConstants.DOC_BI, AppConstants.ENTITY_SUPPLIER, AppConstants.DEFAULT_SUPPLIER_ID,
                                bi_num, float(total_bi), float(total_bi), AppConstants.LOC_STORE, user_name, now_str,
                            ),
                        )
                        new_bi_id = cur.lastrowid
                        cur.execute(
                            '''
                            INSERT INTO transaction_items (
                                transaction_id, product_id, product_name, qty, price, cost_at_time, tva
                            ) VALUES (?, ?, ?, ?, ?, ?, 0)
                            ''',
                            (new_bi_id, p_id, name, float(raw_stock), float(cost), float(cost)),
                        )
                    self.recalculate_stock_for_product(p_id, cur)

            conn.commit()
            return p_id
        except sqlite3.IntegrityError as ie:
            conn.rollback()
            raise ValueError('Un produit portant ce nom existe déjà.') from ie
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def delete_product(self, product_id: int):
        if int(product_id) == AppConstants.DEFAULT_PRODUCT_ID:
            raise ValueError('Impossible de supprimer le produit par défaut.')
        if self.is_product_in_real_use(product_id):
            raise ValueError('Impossible de supprimer : ce produit est utilisé dans des opérations.')
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            cur.execute('BEGIN')
            bi_rows = cur.execute(
                '''
                SELECT DISTINCT t.id FROM transactions t
                JOIN transaction_items ti ON t.id = ti.transaction_id
                WHERE ti.product_id = ? AND t.doc_type = ?
                ''',
                (product_id, AppConstants.DOC_BI),
            ).fetchall()
            for r in bi_rows:
                self._delete_invoice_in_cursor(cur, r['id'])
            cur.execute('DELETE FROM products WHERE id = ?', (product_id,))
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    # --- Tiers (Clients & Fournisseurs) ---
    def get_entities(self, entity_type: str = AppConstants.ENTITY_CLIENT, search_query: str = '', limit: int = 50, offset: int = 0, sort_by: str = 'name', active_ids: list | None = None) -> list:
        conn = self.get_connection()
        try:
            sql = 'SELECT * FROM entities WHERE entity_type = ?'
            params = [entity_type]
            if search_query:
                for term in str(search_query).split():
                    sql += ' AND (name LIKE ? OR phone LIKE ?)'
                    like = f'%{term}%'
                    params.extend([like, like])

            def_id = AppConstants.DEFAULT_CLIENT_ID if entity_type == AppConstants.ENTITY_CLIENT else AppConstants.DEFAULT_SUPPLIER_ID
            if sort_by == 'balance':
                sql += f' ORDER BY CASE WHEN id = {def_id} THEN 0 ELSE 1 END, balance DESC, name COLLATE NOCASE ASC'
            elif sort_by == 'active' and active_ids:
                ids_str = ','.join(str(int(x)) for x in active_ids)
                sql += f' ORDER BY CASE WHEN id = {def_id} THEN 0 WHEN id IN ({ids_str}) THEN 1 ELSE 2 END, name COLLATE NOCASE ASC'
            else:
                sql += f' ORDER BY CASE WHEN id = {def_id} THEN 0 ELSE 1 END, name COLLATE NOCASE ASC'
            sql += ' LIMIT ? OFFSET ?'
            params.extend([limit, offset])
            return [dict(r) for r in conn.execute(sql, params).fetchall()]
        finally:
            conn.close()

    def get_entity_by_id(self, entity_id: int) -> dict | None:
        conn = self.get_connection()
        try:
            row = conn.execute('SELECT * FROM entities WHERE id = ?', (entity_id,)).fetchone()
            return dict(row) if row else None
        finally:
            conn.close()

    def get_active_entity_ids_today(self, entity_type: str = AppConstants.ENTITY_CLIENT) -> list:
        conn = self.get_connection()
        try:
            today_str = datetime.now().strftime('%Y-%m-%d')
            rows = conn.execute(
                '''
                SELECT DISTINCT entity_id FROM transactions
                WHERE date(date) = date(?)
                  AND entity_category = ?
                  AND doc_type NOT IN (?, ?)
                ''',
                (today_str, entity_type, AppConstants.DOC_CLIENT_PAY, AppConstants.DOC_SUPPLIER_PAY),
            ).fetchall()
            return [r['entity_id'] for r in rows if r['entity_id'] is not None]
        finally:
            conn.close()

    def save_entity(self, data: dict) -> int:
        action = data.get('action', 'add')
        e_id = data.get('id')
        e_type = data.get('type', AppConstants.ENTITY_CLIENT)
        name = str(data.get('name', '')).strip()
        if not name:
            raise ValueError('Le nom est obligatoire.')
        if AppConstants.is_default_entity(e_id, e_type):
            raise ValueError('Impossible de modifier le compte par défaut.')
        if name.lower() in (AppConstants.DEFAULT_CLIENT_NAME.lower(), AppConstants.DEFAULT_SUPPLIER_NAME.lower()):
            raise ValueError('Ce nom est réservé au compte par défaut.')

        p_cat = AppConstants.normalize_price_category(data.get('price_category', AppConstants.PRICE_CAT_DETAIL))
        fields = (
            name,
            str(data.get('phone', '')).strip(),
            str(data.get('address', '')).strip(),
            str(data.get('email', '')).strip(),
            str(data.get('activity', '')).strip(),
            str(data.get('rc', '')).strip(),
            str(data.get('nif', '')).strip(),
            str(data.get('nis', '')).strip(),
            str(data.get('nai', '')).strip(),
            str(data.get('rib', '')).strip(),
            p_cat,
            str(data.get('gps_location', '')).strip(),
        )
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            if action == 'add':
                cur.execute(
                    '''
                    INSERT INTO entities (
                        entity_type, name, phone, balance, address, email,
                        activity, rc, nif, nis, nai, rib, price_category, gps_location
                    ) VALUES (?, ?, ?, 0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''',
                    (e_type,) + fields,
                )
                e_id = cur.lastrowid
            else:
                cur.execute(
                    '''
                    UPDATE entities SET
                        name=?, phone=?, address=?, email=?, activity=?,
                        rc=?, nif=?, nis=?, nai=?, rib=?, price_category=?, gps_location=?
                    WHERE id=?
                    ''',
                    fields + (e_id,),
                )
            conn.commit()
            return e_id
        except sqlite3.IntegrityError as ie:
            raise ValueError('Un compte avec ce nom existe déjà.') from ie
        finally:
            conn.close()

    def delete_entity(self, entity_id: int) -> bool:
        if AppConstants.is_default_entity(entity_id):
            raise ValueError('Impossible de supprimer le compte par défaut.')
        conn = self.get_connection()
        try:
            cnt = conn.execute('SELECT COUNT(*) FROM transactions WHERE entity_id = ?', (entity_id,)).fetchone()[0]
            if cnt > 0:
                return False
            conn.execute('DELETE FROM entities WHERE id = ?', (entity_id,))
            conn.commit()
            return True
        finally:
            conn.close()

    # --- Moteur Chronologique de Recalcul Stock & PMP (parité Desktop) ---
    def recalculate_stock_for_product(self, product_id: int, cursor: sqlite3.Cursor):
        if AppConstants.is_default_product(product_id):
            return
        prod_row = cursor.execute(
            'SELECT stock_store, stock_warehouse, purchase_price FROM products WHERE id = ?',
            (product_id,),
        ).fetchone()
        if not prod_row:
            return
        is_unlimited = AppConstants.is_unlimited_stock(prod_row['stock_store']) or AppConstants.is_unlimited_stock(prod_row['stock_warehouse'])
        orig_cost = FinancialEngine.to_decimal(prod_row['purchase_price'])

        inv_rows = cursor.execute(
            '''
            SELECT t.date, t.id, t.doc_type, t.location, ti.qty, ti.price
            FROM transaction_items ti
            JOIN transactions t ON ti.transaction_id = t.id
            WHERE ti.product_id = ?
            ''',
            (product_id,),
        ).fetchall()

        tr_rows = cursor.execute(
            '''
            SELECT st.transfer_date, st.id, st.source_location, st.destination_location, sti.qty
            FROM stock_transfer_items sti
            JOIN stock_transfers st ON sti.transfer_id = st.id
            WHERE sti.product_id = ?
            ''',
            (product_id,),
        ).fetchall()

        timeline = []
        for r in inv_rows:
            timeline.append({
                'date': str(r['date']),
                'order': 0,
                'id': r['id'],
                'kind': 'inv',
                'doc_type': r['doc_type'],
                'loc': r['location'] or AppConstants.LOC_STORE,
                'qty': FinancialEngine.to_decimal(r['qty']),
                'price': FinancialEngine.to_decimal(r['price']),
            })
        for r in tr_rows:
            timeline.append({
                'date': str(r['transfer_date']),
                'order': 1,
                'id': r['id'],
                'kind': 'tr',
                'src': r['source_location'],
                'dst': r['destination_location'],
                'qty': FinancialEngine.to_decimal(r['qty']),
            })
        timeline.sort(key=lambda x: (x['date'], x['order'], x['id']))

        running_global = Decimal('0')
        running_pmp = Decimal('0')
        running_store = Decimal('0')
        running_wh = Decimal('0')
        has_purchases = False
        sale_cost_updates = []

        for ev in timeline:
            qty = ev['qty']
            if ev['kind'] == 'inv':
                dt_code = ev['doc_type']
                direction = AppConstants.STOCK_MOVEMENTS.get(dt_code, 0)
                if direction == 1:
                    if dt_code in (AppConstants.DOC_BA, AppConstants.DOC_FF, AppConstants.DOC_BI):
                        has_purchases = True
                        qty_for_avg = running_global if running_global > 0 else Decimal('0')
                        new_total_qty = qty_for_avg + qty
                        if new_total_qty > 0:
                            running_pmp = (qty_for_avg * running_pmp + qty * ev['price']) / new_total_qty
                    elif dt_code == AppConstants.DOC_RC:
                        sale_cost_updates.append((ev['id'], running_pmp if has_purchases else orig_cost))
                    running_global += qty
                    if ev['loc'] == AppConstants.LOC_WAREHOUSE:
                        running_wh += qty
                    else:
                        running_store += qty
                elif direction == -1:
                    if dt_code in (AppConstants.DOC_BV, AppConstants.DOC_FC):
                        sale_cost_updates.append((ev['id'], running_pmp if has_purchases else orig_cost))
                    running_global -= qty
                    if ev['loc'] == AppConstants.LOC_WAREHOUSE:
                        running_wh -= qty
                    else:
                        running_store -= qty
            else:
                if ev['src'] == AppConstants.LOC_STORE:
                    running_store -= qty
                elif ev['src'] == AppConstants.LOC_WAREHOUSE:
                    running_wh -= qty
                if ev['dst'] == AppConstants.LOC_STORE:
                    running_store += qty
                elif ev['dst'] == AppConstants.LOC_WAREHOUSE:
                    running_wh += qty

        final_pmp = FinancialEngine.quantize_money(running_pmp) if has_purchases else orig_cost
        if abs(final_pmp) >= AppConstants.MAX_AMOUNT_LIMIT:
            final_pmp = Decimal('0.00')

        if is_unlimited:
            final_store = Decimal(str(AppConstants.UNLIMITED_STOCK_VALUE))
            final_wh = Decimal(str(AppConstants.UNLIMITED_STOCK_VALUE))
        else:
            min_safe = Decimal(str(AppConstants.UNLIMITED_STOCK_THRESHOLD)) + Decimal('1')
            final_store = max(min(FinancialEngine.quantize_qty(running_store), AppConstants.MAX_STOCK_DB_LIMIT), min_safe)
            final_wh = max(min(FinancialEngine.quantize_qty(running_wh), AppConstants.MAX_STOCK_DB_LIMIT), min_safe)

        cursor.execute(
            'UPDATE products SET stock_store = ?, stock_warehouse = ?, purchase_price = ? WHERE id = ?',
            (float(final_store), float(final_wh), float(final_pmp), product_id),
        )
        for tx_id_item, cost_at_ev in sale_cost_updates:
            eff_cost = FinancialEngine.quantize_money(cost_at_ev)
            if eff_cost <= Decimal('0') and final_pmp > Decimal('0'):
                eff_cost = final_pmp
            if abs(eff_cost) >= AppConstants.MAX_AMOUNT_LIMIT:
                eff_cost = Decimal('0.00')
            cursor.execute(
                '''
                UPDATE transaction_items
                SET cost_at_time = ?
                WHERE transaction_id = ? AND product_id = ?
                ''',
                (float(eff_cost), tx_id_item, product_id),
            )

    def _reverse_invoice_effects(self, cur: sqlite3.Cursor, tx_id: int) -> set:
        old_tx = cur.execute('SELECT * FROM transactions WHERE id = ?', (tx_id,)).fetchone()
        if not old_tx:
            return set()
        old_tx = dict(old_tx)
        old_eid = old_tx['entity_id']
        old_cat = old_tx['entity_category']
        old_doc = old_tx['doc_type']
        old_total = FinancialEngine.to_decimal(old_tx['total_amount'])
        old_paid = FinancialEngine.to_decimal(old_tx['paid_amount'])

        affected_pids = set()
        for row in cur.execute('SELECT product_id FROM transaction_items WHERE transaction_id = ?', (tx_id,)).fetchall():
            if not AppConstants.is_default_product(row['product_id']):
                affected_pids.add(row['product_id'])

        if old_eid and not AppConstants.is_default_entity(old_eid, old_cat):
            f_factor = AppConstants.FINANCIAL_FACTORS.get(old_doc, 0)
            s_factor = AppConstants.STOCK_MOVEMENTS.get(old_doc, 0)
            if f_factor == -1 and s_factor == 0:
                old_impact = -old_total
            elif f_factor == -1 and s_factor != 0:
                old_impact = -old_total + old_paid
            else:
                old_impact = (old_total * Decimal(f_factor)) - old_paid
            if old_impact != Decimal('0'):
                cur.execute(
                    'UPDATE entities SET balance = ROUND(balance - ?, 2) WHERE id = ?',
                    (float(old_impact), old_eid),
                )

        cur.execute('DELETE FROM transaction_items WHERE transaction_id = ?', (tx_id,))
        return affected_pids

    def _delete_invoice_in_cursor(self, cur: sqlite3.Cursor, tx_id: int):
        affected_pids = self._reverse_invoice_effects(cur, tx_id)
        cur.execute('DELETE FROM transactions WHERE id = ?', (tx_id,))
        for pid in affected_pids:
            self.recalculate_stock_for_product(pid, cur)

    def save_transaction(self, tx_data: dict) -> int:
        tx_id = tx_data.get('id')
        doc_type = tx_data['doc_type']
        entity_category = AppConstants.get_entity_type(doc_type)
        entity_id = tx_data.get('entity_id')
        if entity_id is None:
            entity_id = AppConstants.DEFAULT_SUPPLIER_ID if entity_category == AppConstants.ENTITY_SUPPLIER else AppConstants.DEFAULT_CLIENT_ID

        is_simple = bool(tx_data.get('is_simple_payment', False))
        items = tx_data.get('items', [])
        location = tx_data.get('location', AppConstants.LOC_STORE)
        note = str(tx_data.get('note', '')).strip()
        user_name = str(tx_data.get('user_name', AppConstants.ADMIN_USERNAME))
        date_str = str(tx_data.get('timestamp') or datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        global_discount = FinancialEngine.quantize_money(tx_data.get('global_discount', 0))
        payment_method = str(tx_data.get('payment_method', AppConstants.PAY_DEFAULT))

        if not is_simple and not items:
            raise ValueError('Le panier est vide.')

        if doc_type == AppConstants.DOC_BI:
            if len(items) != 1:
                raise ValueError('Le Bon Initial (BI) doit contenir exactement un seul produit.')
            if AppConstants.is_default_product(items[0]):
                raise ValueError("Action interdite sur le produit système dans un Bon Initial.")

        if doc_type == AppConstants.DOC_TR:
            return self.save_stock_transfer(
                transfer_id=tx_id,
                date_str=date_str,
                source_loc=location,
                dest_loc=AppConstants.LOC_WAREHOUSE if location == AppConstants.LOC_STORE else AppConstants.LOC_STORE,
                notes=note,
                items=items,
                user_name=user_name,
            )

        is_comptoir = AppConstants.is_default_entity(entity_id, entity_category)
        if is_simple and is_comptoir:
            raise ValueError('Impossible de créer un versement/crédit pour le compte Comptoir.')

        conn = self.get_connection()
        try:
            cur = conn.cursor()
            cur.execute('BEGIN')

            if doc_type == AppConstants.DOC_BI and tx_id:
                orig_bi_item = cur.execute(
                    'SELECT product_id FROM transaction_items WHERE transaction_id = ? LIMIT 1',
                    (tx_id,),
                ).fetchone()
                if orig_bi_item and int(orig_bi_item['product_id']) != int(items[0]['id']):
                    raise ValueError('Impossible de remplacer le produit d\'un Bon Initial (BI).')

            s_factor = AppConstants.STOCK_MOVEMENTS.get(doc_type, 0)
            allow_negative = self.get_bool_setting('allow_negative_stock', default=True)
            if not is_simple and s_factor == -1 and not allow_negative:
                old_loc_qtys = {}
                if tx_id:
                    old_tx_chk = cur.execute('SELECT doc_type, location FROM transactions WHERE id = ?', (tx_id,)).fetchone()
                    if old_tx_chk and (old_tx_chk['location'] or AppConstants.LOC_STORE) == location:
                        old_dir = AppConstants.STOCK_MOVEMENTS.get(old_tx_chk['doc_type'], 0)
                        if old_dir != 0:
                            for r in cur.execute('SELECT product_id, qty FROM transaction_items WHERE transaction_id = ?', (tx_id,)).fetchall():
                                if not AppConstants.is_default_product(r['product_id']):
                                    old_loc_qtys[r['product_id']] = old_loc_qtys.get(r['product_id'], Decimal('0')) + (FinancialEngine.to_decimal(r['qty']) * Decimal(old_dir))

                needed_by_pid = {}
                for it in items:
                    pid = int(it['id'])
                    if AppConstants.is_default_product(pid):
                        continue
                    needed_by_pid[pid] = needed_by_pid.get(pid, Decimal('0')) + FinancialEngine.to_decimal(it['qty'])

                col = 'stock_warehouse' if location == AppConstants.LOC_WAREHOUSE else 'stock_store'
                for pid, req_qty in needed_by_pid.items():
                    row_p = cur.execute(f'SELECT {col}, name FROM products WHERE id = ?', (pid,)).fetchone()
                    if row_p:
                        avail = FinancialEngine.to_decimal(row_p[col])
                        if not AppConstants.is_unlimited_stock(avail):
                            effective_avail = avail - old_loc_qtys.get(pid, Decimal('0'))
                            if effective_avail < req_qty:
                                loc_lbl = AppConstants.LOC_LABELS.get(location, 'Magasin')
                                raise ValueError(
                                    f"Stock insuffisant ({loc_lbl}) pour '{row_p['name']}'.\n"
                                    f'Disponible : {FinancialEngine.format_qty(effective_avail)} | Demandé : {FinancialEngine.format_qty(req_qty)}'
                                )

            if is_simple:
                final_total = FinancialEngine.quantize_money(tx_data['amount'])
                paid_amount = final_total
                stamp_duty = Decimal('0.00')
                doc_number = tx_data.get('doc_number') or ('Règlement' if doc_type == AppConstants.DOC_SUPPLIER_PAY else 'Versement')
                if final_total < 0:
                    doc_number = 'Dette Initiale' if entity_category == AppConstants.ENTITY_SUPPLIER else 'Crédit Initial'
            else:
                calc = FinancialEngine.calculate_document_totals(
                    items=items,
                    doc_type=doc_type,
                    global_discount=global_discount,
                    payment_method=payment_method,
                    is_comptoir=is_comptoir,
                )
                final_total = calc['final_total']
                stamp_duty = calc['stamp_duty']
                paid_amount = FinancialEngine.quantize_money(tx_data.get('paid_amount', 0))
                if is_comptoir and doc_type not in AppConstants.NON_FINANCIAL_DOC_TYPES:
                    paid_amount = final_total
                doc_number = tx_data.get('doc_number') or ''
                if tx_id and not doc_number:
                    r_num = cur.execute('SELECT doc_number FROM transactions WHERE id = ?', (tx_id,)).fetchone()
                    doc_number = r_num['doc_number'] if r_num else self.peek_next_doc_number(doc_type, cur)
                elif not doc_number:
                    doc_number = self.peek_next_doc_number(doc_type, cur)

            ok_bounds, err_bounds = FinancialEngine.validate_transaction_bounds(final_total, paid_amount)
            if not ok_bounds:
                raise ValueError(err_bounds)

            affected_pids = set()
            if tx_id:
                affected_pids = self._reverse_invoice_effects(cur, tx_id)
                cur.execute(
                    '''
                    UPDATE transactions SET
                        doc_type=?, entity_category=?, entity_id=?, doc_number=?,
                        total_amount=?, global_discount=?, paid_amount=?, payment_method=?,
                        stamp_duty=?, location=?, note=?, user_name=?
                    WHERE id=?
                    ''',
                    (
                        doc_type, entity_category, entity_id, doc_number,
                        float(final_total), float(global_discount), float(paid_amount), payment_method,
                        float(stamp_duty), location, note, user_name, tx_id,
                    ),
                )
            else:
                cur.execute(
                    '''
                    INSERT INTO transactions (
                        doc_type, entity_category, entity_id, doc_number,
                        total_amount, global_discount, paid_amount, payment_method,
                        stamp_duty, location, note, user_name, date
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''',
                    (
                        doc_type, entity_category, entity_id, doc_number,
                        float(final_total), float(global_discount), float(paid_amount), payment_method,
                        float(stamp_duty), location, note, user_name, date_str,
                    ),
                )
                tx_id = cur.lastrowid

            for it in items:
                pid = int(it['id'])
                qty = FinancialEngine.quantize_qty(it['qty'])
                price = FinancialEngine.quantize_money(it['price'])
                tva = FinancialEngine.quantize_money(it.get('tva', 0)) if doc_type in AppConstants.INVOICE_TAX_DOC_TYPES else Decimal('0.00')
                d_amt = FinancialEngine.quantize_money(it.get('discount_amount', 0))
                d_type = it.get('discount_type', 'fixed')
                p_name = str(it.get('name', 'Article'))

                cost_at_time = Decimal('0.00')
                if not AppConstants.is_default_product(pid):
                    affected_pids.add(pid)
                    if doc_type == AppConstants.DOC_BI:
                        cost_at_time = price
                    else:
                        p_row = cur.execute('SELECT purchase_price FROM products WHERE id = ?', (pid,)).fetchone()
                        if p_row:
                            cost_at_time = FinancialEngine.quantize_money(p_row['purchase_price'])

                cur.execute(
                    '''
                    INSERT INTO transaction_items (
                        transaction_id, product_id, product_name, qty, price,
                        cost_at_time, tva, discount_amount, discount_type
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''',
                    (
                        tx_id, pid, p_name, float(qty), float(price),
                        float(cost_at_time), float(tva), float(d_amt), d_type,
                    ),
                )

            for pid in affected_pids:
                self.recalculate_stock_for_product(pid, cur)

            if not is_comptoir:
                f_factor = AppConstants.FINANCIAL_FACTORS.get(doc_type, 0)
                if f_factor == -1 and s_factor == 0:
                    bal_impact = -final_total
                elif f_factor == -1 and s_factor != 0:
                    bal_impact = -final_total + paid_amount
                else:
                    bal_impact = (final_total * Decimal(f_factor)) - paid_amount
                if bal_impact != Decimal('0'):
                    cur.execute(
                        'UPDATE entities SET balance = ROUND(balance + ?, 2) WHERE id = ?',
                        (float(bal_impact), entity_id),
                    )

            conn.commit()
            return tx_id
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def save_stock_transfer(self, transfer_id, date_str: str, source_loc: str, dest_loc: str, notes: str, items: list, user_name: str) -> int:
        if source_loc == dest_loc:
            raise ValueError('La source et la destination doivent être différentes.')
        if not items:
            raise ValueError('Le panier de transfert est vide.')

        conn = self.get_connection()
        try:
            cur = conn.cursor()
            cur.execute('BEGIN')

            old_qtys_by_pid = {}
            if transfer_id:
                old_tr = cur.execute('SELECT source_location, destination_location FROM stock_transfers WHERE id = ?', (transfer_id,)).fetchone()
                if old_tr:
                    for r in cur.execute('SELECT product_id, qty FROM stock_transfer_items WHERE transfer_id = ?', (transfer_id,)).fetchall():
                        pid = r['product_id']
                        q = FinancialEngine.to_decimal(r['qty'])
                        if old_tr['source_location'] == source_loc:
                            old_qtys_by_pid[pid] = old_qtys_by_pid.get(pid, Decimal('0')) + q
                        elif old_tr['destination_location'] == source_loc:
                            old_qtys_by_pid[pid] = old_qtys_by_pid.get(pid, Decimal('0')) - q

            req_by_pid = {}
            for it in items:
                pid = int(it['id'])
                if AppConstants.is_default_product(pid):
                    raise ValueError("Impossible de transférer l'article système.")
                qty = FinancialEngine.quantize_qty(it['qty'])
                if qty <= 0 or qty > AppConstants.MAX_QTY_LIMIT:
                    raise ValueError('Quantité de transfert invalide.')
                req_by_pid[pid] = req_by_pid.get(pid, Decimal('0')) + qty

            src_col = 'stock_warehouse' if source_loc == AppConstants.LOC_WAREHOUSE else 'stock_store'
            for pid, req_qty in req_by_pid.items():
                p_row = cur.execute(f'SELECT {src_col}, stock_store, stock_warehouse, name FROM products WHERE id = ?', (pid,)).fetchone()
                if not p_row:
                    raise ValueError('Produit introuvable.')
                if AppConstants.is_unlimited_stock(p_row['stock_store']) or AppConstants.is_unlimited_stock(p_row['stock_warehouse']):
                    raise ValueError(f"Impossible de transférer un produit à stock illimité ({p_row['name']}).")
                avail = FinancialEngine.to_decimal(p_row[src_col]) + old_qtys_by_pid.get(pid, Decimal('0'))
                if avail < req_qty:
                    loc_lbl = AppConstants.LOC_LABELS.get(source_loc, 'Magasin')
                    raise ValueError(
                        f"Stock insuffisant ({loc_lbl}) pour '{p_row['name']}'.\n"
                        f'Disponible : {FinancialEngine.format_qty(avail)} | Demandé : {FinancialEngine.format_qty(req_qty)}'
                    )

            affected_pids = set(req_by_pid.keys())
            if transfer_id:
                for r in cur.execute('SELECT product_id FROM stock_transfer_items WHERE transfer_id = ?', (transfer_id,)).fetchall():
                    affected_pids.add(r['product_id'])
                cur.execute('DELETE FROM stock_transfer_items WHERE transfer_id = ?', (transfer_id,))
                cur.execute(
                    'UPDATE stock_transfers SET transfer_date=?, source_location=?, destination_location=?, notes=?, user_name=? WHERE id=?',
                    (date_str, source_loc, dest_loc, notes, user_name, transfer_id),
                )
            else:
                tr_num = self.peek_next_doc_number(AppConstants.DOC_TR, cur)
                cur.execute(
                    'INSERT INTO stock_transfers (transfer_number, transfer_date, source_location, destination_location, notes, user_name) VALUES (?, ?, ?, ?, ?, ?)',
                    (tr_num, date_str, source_loc, dest_loc, notes, user_name),
                )
                transfer_id = cur.lastrowid

            for it in items:
                cur.execute(
                    'INSERT INTO stock_transfer_items (transfer_id, product_id, product_name, qty) VALUES (?, ?, ?, ?)',
                    (transfer_id, int(it['id']), str(it['name']), float(FinancialEngine.quantize_qty(it['qty']))),
                )

            for pid in affected_pids:
                self.recalculate_stock_for_product(pid, cur)

            conn.commit()
            return transfer_id
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def delete_transaction(self, tx_id: int, is_transfer: bool = False):
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            cur.execute('BEGIN')
            if is_transfer:
                affected = {
                    r['product_id']
                    for r in cur.execute('SELECT product_id FROM stock_transfer_items WHERE transfer_id = ?', (tx_id,)).fetchall()
                }
                cur.execute('DELETE FROM stock_transfers WHERE id = ?', (tx_id,))
                for pid in affected:
                    self.recalculate_stock_for_product(pid, cur)
            else:
                self._delete_invoice_in_cursor(cur, tx_id)
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    # --- Dépenses / Charges ---
    def save_expense(self, description: str, amount, category: str, recorded_by: str, expense_id: int | None = None, date_str: str | None = None) -> int:
        amt = FinancialEngine.quantize_money(amount)
        if amt <= Decimal('0') or amt >= AppConstants.MAX_AMOUNT_LIMIT:
            raise ValueError('Montant de dépense invalide.')
        ts = date_str or datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            if expense_id:
                cur.execute(
                    'UPDATE general_expenses SET description=?, amount=?, category=? WHERE id=?',
                    (description, float(amt), category, expense_id),
                )
                res_id = expense_id
            else:
                cur.execute(
                    'INSERT INTO general_expenses (description, amount, category, expense_date, recorded_by) VALUES (?, ?, ?, ?, ?)',
                    (description, float(amt), category, ts, recorded_by),
                )
                res_id = cur.lastrowid
            conn.commit()
            return res_id
        finally:
            conn.close()

    def get_expenses(self, start_date: str, end_date: str) -> list:
        conn = self.get_connection()
        try:
            rows = conn.execute(
                '''
                SELECT * FROM general_expenses
                WHERE date(expense_date) BETWEEN date(?) AND date(?)
                ORDER BY expense_date DESC, id DESC
                ''',
                (start_date, end_date),
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()

    def delete_expense(self, expense_id: int):
        conn = self.get_connection()
        try:
            conn.execute('DELETE FROM general_expenses WHERE id = ?', (expense_id,))
            conn.commit()
        finally:
            conn.close()

    # --- Historique & Détails ---
    def get_operations_history(
        self,
        start_date: str,
        end_date: str,
        entity_id: int | None = None,
        entity_category: str | None = None,
        doc_type_filter: str = AppConstants.FAM_ALL,
        limit: int = 50,
        offset: int = 0,
    ) -> list:
        is_seller_only = doc_type_filter == 'SELLER_ONLY'
        include_tx = doc_type_filter != AppConstants.DOC_TR
        include_tr = (
            not is_seller_only
            and entity_id is None
            and not entity_category
            and doc_type_filter in (AppConstants.FAM_ALL, AppConstants.DOC_TR)
        )
        if not include_tx and not include_tr:
            return []

        conn = self.get_connection()
        try:
            branches = []
            params = []

            if include_tx:
                tx_sql = '''
                    SELECT t.id, t.doc_type, t.entity_category, t.entity_id, t.doc_number,
                           t.total_amount, t.global_discount, t.paid_amount, t.payment_method,
                           t.stamp_duty, t.location, t.note, t.user_name, t.date,
                           e.name AS entity_name,
                           '' AS source_location,
                           '' AS destination_location,
                           0 AS is_transfer
                    FROM transactions t
                    LEFT JOIN entities e ON t.entity_id = e.id
                    WHERE date(t.date) BETWEEN date(?) AND date(?)
                '''
                params.extend([start_date, end_date])
                if entity_id is not None:
                    tx_sql += ' AND t.entity_id = ?'
                    params.append(entity_id)
                if entity_category:
                    tx_sql += ' AND t.entity_category = ?'
                    params.append(entity_category)
                if is_seller_only:
                    tx_sql += ' AND t.doc_type IN (?, ?, ?)'
                    params.extend([AppConstants.DOC_BV, AppConstants.DOC_RC, AppConstants.DOC_CLIENT_PAY])
                elif doc_type_filter != AppConstants.FAM_ALL:
                    tx_sql += ' AND t.doc_type = ?'
                    params.append(doc_type_filter)
                branches.append(tx_sql)

            if include_tr:
                tr_sql = '''
                    SELECT st.id, ? AS doc_type, ? AS entity_category, NULL AS entity_id,
                           st.transfer_number AS doc_number,
                           0.0 AS total_amount, 0.0 AS global_discount, 0.0 AS paid_amount,
                           '' AS payment_method, 0.0 AS stamp_duty, st.source_location AS location,
                           st.notes AS note, st.user_name, st.transfer_date AS date,
                           '' AS entity_name,
                           st.source_location,
                           st.destination_location,
                           1 AS is_transfer
                    FROM stock_transfers st
                    WHERE date(st.transfer_date) BETWEEN date(?) AND date(?)
                '''
                params.extend([AppConstants.DOC_TR, AppConstants.ENTITY_CLIENT, start_date, end_date])
                branches.append(tr_sql)

            union_sql = ' UNION ALL '.join(branches)
            full_sql = f'SELECT * FROM ({union_sql}) ORDER BY date DESC, id DESC LIMIT ? OFFSET ?'
            params.extend([limit, offset])

            out = []
            for r in conn.execute(full_sql, params).fetchall():
                d = dict(r)
                d['is_transfer'] = bool(d.get('is_transfer', 0))
                out.append(d)
            return out
        finally:
            conn.close()

    def get_operation_details(self, op_id: int, is_transfer: bool = False) -> dict | None:
        conn = self.get_connection()
        try:
            if is_transfer:
                tr = conn.execute('SELECT * FROM stock_transfers WHERE id = ?', (op_id,)).fetchone()
                if not tr:
                    return None
                tr_d = dict(tr)
                tr_d['doc_type'] = AppConstants.DOC_TR
                tr_d['doc_number'] = tr_d['transfer_number']
                tr_d['date'] = tr_d['transfer_date']
                tr_d['location'] = tr_d['source_location']
                tr_d['is_transfer'] = True
                items = [
                    dict(r)
                    for r in conn.execute(
                        '''
                        SELECT sti.*, p.product_ref, p.barcode
                        FROM stock_transfer_items sti
                        LEFT JOIN products p ON sti.product_id = p.id
                        WHERE sti.transfer_id = ?
                        ''',
                        (op_id,),
                    ).fetchall()
                ]
                return {'transaction': tr_d, 'entity': {}, 'items': items}

            tx = conn.execute('SELECT * FROM transactions WHERE id = ?', (op_id,)).fetchone()
            if not tx:
                return None
            tx_d = dict(tx)
            tx_d['is_transfer'] = False
            ent = {}
            if tx_d.get('entity_id'):
                e_row = conn.execute('SELECT * FROM entities WHERE id = ?', (tx_d['entity_id'],)).fetchone()
                if e_row:
                    ent = dict(e_row)
            items = [
                dict(r)
                for r in conn.execute(
                    '''
                    SELECT ti.*, p.product_ref, p.barcode
                    FROM transaction_items ti
                    LEFT JOIN products p ON ti.product_id = p.id
                    WHERE ti.transaction_id = ?
                    ''',
                    (op_id,),
                ).fetchall()
            ]
            return {'transaction': tx_d, 'entity': ent, 'items': items}
        finally:
            conn.close()

    # --- Statistiques Financières (100% parité Desktop) ---
    def get_comprehensive_stats(self, start_date: str, end_date: str) -> dict:
        conn = self.get_connection()
        stats = {
            'sales': Decimal('0.00'),
            'purchases': Decimal('0.00'),
            'gross_profit': Decimal('0.00'),
            'expenses': Decimal('0.00'),
            'profit': Decimal('0.00'),
            'cash_in': Decimal('0.00'),
            'cash_out': Decimal('0.00'),
            'caisse_total': Decimal('0.00'),
            'stock_value_cost': Decimal('0.00'),
            'stock_value_sale': Decimal('0.00'),
            'client_debts': Decimal('0.00'),
            'supplier_dues': Decimal('0.00'),
        }
        ventes_payees = Decimal('0.00')
        versements_clients = Decimal('0.00')
        rc_comptoir = Decimal('0.00')
        try:
            cur = conn.cursor()
            c_row = cur.execute(
                'SELECT COALESCE(SUM(balance), 0) FROM entities WHERE entity_type = ? AND id <> ?',
                (AppConstants.ENTITY_CLIENT, AppConstants.DEFAULT_CLIENT_ID),
            ).fetchone()
            stats['client_debts'] = FinancialEngine.to_decimal(c_row[0])

            s_row = cur.execute(
                'SELECT COALESCE(SUM(balance), 0) FROM entities WHERE entity_type = ? AND id <> ?',
                (AppConstants.ENTITY_SUPPLIER, AppConstants.DEFAULT_SUPPLIER_ID),
            ).fetchone()
            stats['supplier_dues'] = FinancialEngine.to_decimal(s_row[0])

            for p in cur.execute(
                'SELECT stock_store, stock_warehouse, purchase_price, price FROM products WHERE id <> ?',
                (AppConstants.DEFAULT_PRODUCT_ID,),
            ).fetchall():
                s_st = FinancialEngine.to_decimal(p['stock_store'])
                s_wh = FinancialEngine.to_decimal(p['stock_warehouse'])
                if AppConstants.is_unlimited_stock(s_st) or AppConstants.is_unlimited_stock(s_wh):
                    continue
                tot_qty = max(Decimal('0'), s_st) + max(Decimal('0'), s_wh)
                if tot_qty > 0:
                    stats['stock_value_cost'] += tot_qty * FinancialEngine.to_decimal(p['purchase_price'])
                    stats['stock_value_sale'] += tot_qty * FinancialEngine.to_decimal(p['price'])

            exp_row = cur.execute(
                'SELECT COALESCE(SUM(amount), 0) FROM general_expenses WHERE date(expense_date) BETWEEN date(?) AND date(?)',
                (start_date, end_date),
            ).fetchone()
            stats['expenses'] = FinancialEngine.to_decimal(exp_row[0])

            tx_rows = cur.execute(
                'SELECT * FROM transactions WHERE date(date) BETWEEN date(?) AND date(?)',
                (start_date, end_date),
            ).fetchall()

            for tx in tx_rows:
                dt_code = tx['doc_type']
                total_amt = FinancialEngine.to_decimal(tx['total_amount'])
                paid_amt = FinancialEngine.to_decimal(tx['paid_amount'])
                g_disc = FinancialEngine.to_decimal(tx['global_discount'])

                if dt_code == AppConstants.DOC_CLIENT_PAY:
                    if total_amt > Decimal('0'):
                        stats['cash_in'] += total_amt
                        versements_clients += total_amt
                    continue
                if dt_code == AppConstants.DOC_SUPPLIER_PAY:
                    if total_amt > Decimal('0'):
                        stats['cash_out'] += total_amt
                    continue

                if dt_code in (AppConstants.DOC_BV, AppConstants.DOC_FC):
                    stats['sales'] += total_amt
                    stats['cash_in'] += paid_amt
                    ventes_payees += paid_amt
                elif dt_code == AppConstants.DOC_RC:
                    stats['sales'] -= total_amt
                    stats['cash_out'] += paid_amt
                    if AppConstants.is_default_entity(tx['entity_id'], tx['entity_category']):
                        rc_comptoir += abs(total_amt)
                elif dt_code in (AppConstants.DOC_BA, AppConstants.DOC_FF, AppConstants.DOC_BI):
                    stats['purchases'] += total_amt
                    if dt_code != AppConstants.DOC_BI:
                        stats['cash_out'] += paid_amt
                elif dt_code == AppConstants.DOC_RF:
                    stats['purchases'] -= total_amt
                    stats['cash_in'] += paid_amt

                if dt_code in (AppConstants.DOC_BV, AppConstants.DOC_FC, AppConstants.DOC_RC):
                    items_rows = [
                        dict(r)
                        for r in cur.execute(
                            'SELECT * FROM transaction_items WHERE transaction_id = ?',
                            (tx['id'],),
                        ).fetchall()
                    ]
                    if items_rows:
                        calc = FinancialEngine.calculate_document_totals(
                            items_rows,
                            doc_type=dt_code,
                            global_discount=g_disc,
                            payment_method=tx['payment_method'] or '',
                            is_comptoir=AppConstants.is_default_entity(tx['entity_id'], tx['entity_category']),
                        )
                        sign = Decimal('-1') if dt_code == AppConstants.DOC_RC else Decimal('1')
                        real_ht = Decimal('0.00')
                        all_ht = Decimal('0.00')
                        cogs = Decimal('0.00')
                        has_real_items = False
                        for it in items_rows:
                            q = FinancialEngine.to_decimal(it['qty'])
                            p = FinancialEngine.to_decimal(it['price'])
                            d_amt = FinancialEngine.to_decimal(it['discount_amount'])
                            u_disc = d_amt if it['discount_type'] == 'fixed' else p * (d_amt / Decimal('100'))
                            line_ht = FinancialEngine.quantize_money((p - u_disc) * q)
                            all_ht += line_ht
                            if not AppConstants.is_default_product(it['product_id']):
                                has_real_items = True
                                real_ht += line_ht
                                cogs += FinancialEngine.quantize_money(q * FinancialEngine.to_decimal(it['cost_at_time']))
                        if has_real_items:
                            gd_ht = calc['total_ht_net_of_lines'] - calc['total_ht_final']
                            if all_ht > Decimal('0') and real_ht < all_ht:
                                gd_ht = FinancialEngine.quantize_money(gd_ht * (real_ht / all_ht))
                            elif real_ht == Decimal('0'):
                                gd_ht = Decimal('0.00')
                            stats['gross_profit'] += (real_ht - gd_ht - cogs) * sign

            stats['profit'] = stats['gross_profit'] - stats['expenses']
            stats['cash_out'] += stats['expenses']
            stats['caisse_total'] = ventes_payees + versements_clients - rc_comptoir - stats['expenses']
        finally:
            conn.close()

        return {k: FinancialEngine.quantize_money(v) for k, v in stats.items()}


# ==============================================================================
# 5. GÉNÉRATEUR PDF A4 & IMPRESSION THERMIQUE BLUETOOTH 80 MM
# ==============================================================================
class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        logging.getLogger('fontTools').setLevel(logging.ERROR)
        super().__init__(*args, **kwargs)
        self.store_info = {}
        self.entity_info = {}
        self.doc_info = {}
        self.totals = {}
        self.payment_info = {}
        self.amount_in_words = ''
        self.balance_data = None
        self.logo_path = ''
        try:
            if os.path.exists(FONT_FILE):
                self.add_font('arabic', '', FONT_FILE)
                self.add_font('arabic', 'B', FONT_FILE)
            else:
                self.add_font('arabic', '')
                self.add_font('arabic', 'B', '')
        except Exception:
            pass

    def smart_text(self, text) -> str:
        return ar_visual(str(text) if text is not None else '')

    def _draw_logo_box(self, x: float, y: float, max_w: float, max_h: float, center_x: bool = False) -> float:
        lp = str(self.logo_path or '').strip()
        if not lp or not os.path.exists(lp):
            return 0.0
        try:
            with Image.open(lp) as im:
                iw, ih = im.size
            if iw <= 0 or ih <= 0:
                return 0.0
            scale = min(max_w / float(iw), max_h / float(ih))
            dw = float(iw) * scale
            dh = float(ih) * scale
            draw_x = (self.w - dw) / 2.0 if center_x else x
            draw_y = y if center_x else y + (max_h - dh) / 2.0
            self.image(lp, x=draw_x, y=draw_y, w=dw, h=dh)
            return dh
        except Exception:
            return 0.0

    def header(self):
        if self.page_no() != 1:
            return
        store_name = self.smart_text(self.store_info.get('store_name', AppConstants.APP_NAME))
        doc_type = self.doc_info.get('doc_type')
        if doc_type in (AppConstants.DOC_CLIENT_PAY, AppConstants.DOC_SUPPLIER_PAY):
            dh = self._draw_logo_box(0, 8, 32, 16, center_x=True)
            self.set_y(10 + (dh + 2 if dh > 0 else 0))
            self.set_font('arabic', 'B', 20)
            self.cell(0, 10, store_name, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
            self.ln(4)
            return

        info_lines = []
        if (act := self.smart_text(self.store_info.get('store_activity'))):
            info_lines.append(act)
        if (addr := self.smart_text(self.store_info.get('store_address'))):
            info_lines.append(addr)
        pe = ' / '.join(
            filter(
                None,
                [
                    f"Tél : {self.store_info.get('store_phone')}" if self.store_info.get('store_phone') else None,
                    f"Email : {self.store_info.get('store_email')}" if self.store_info.get('store_email') else None,
                ],
            )
        )
        if pe:
            info_lines.append(pe)

        if doc_type not in (AppConstants.DOC_BV, AppConstants.DOC_BA):
            rn = ' / '.join(
                filter(
                    None,
                    [
                        f"RC : {self.store_info.get('store_rc')}" if self.store_info.get('store_rc') else None,
                        f"NIF : {self.store_info.get('store_nif')}" if self.store_info.get('store_nif') else None,
                    ],
                )
            )
            if rn:
                info_lines.append(rn)
            nn = ' / '.join(
                filter(
                    None,
                    [
                        f"N.Art : {self.store_info.get('store_nai')}" if self.store_info.get('store_nai') else None,
                        f"NIS : {self.store_info.get('store_nis')}" if self.store_info.get('store_nis') else None,
                    ],
                )
            )
            if nn:
                info_lines.append(nn)
            if (rib := self.smart_text(self.store_info.get('store_rib'))):
                info_lines.append(f'RIB : {rib}')

        self._draw_logo_box(10, 8, 30, 24, center_x=False)
        self.set_y(8)
        block_w = 120
        self.set_x((self.w - block_w) / 2)
        self.set_font('arabic', 'B', 22)
        self.cell(block_w, 10, store_name, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.set_x((self.w - block_w) / 2)
        self.set_font('arabic', '', 9)
        self.multi_cell(block_w, 5, '\n'.join(info_lines), align='C')
        self.set_y(max(self.get_y(), 34))
        self.ln(6)

        y_start = self.get_y()
        self.set_y(y_start)
        self.set_x(-100)
        self.set_font('arabic', 'B', 10)
        self.multi_cell(
            90,
            7,
            f"Date : {self.doc_info.get('date', '')}\n{self.doc_info.get('doc_name_fr', '')}\nN° : {self.doc_info.get('doc_number', '')}",
            border=1,
            align='L',
        )
        y_inv_end = self.get_y()

        self.set_y(y_start)
        self.set_x(10)
        c_lines = []
        if (ca := self.smart_text(self.entity_info.get('address'))):
            c_lines.append(f'Adresse : {ca}')
        cpe = ' / '.join(
            filter(
                None,
                [
                    f"Tél : {self.entity_info.get('phone')}" if self.entity_info.get('phone') else None,
                    f"Email : {self.entity_info.get('email')}" if self.entity_info.get('email') else None,
                ],
            )
        )
        if cpe:
            c_lines.append(cpe)
        crn = ' / '.join(
            filter(
                None,
                [
                    f"RC : {self.entity_info.get('rc')}" if self.entity_info.get('rc') else None,
                    f"NIF : {self.entity_info.get('nif')}" if self.entity_info.get('nif') else None,
                ],
            )
        )
        if crn:
            c_lines.append(crn)
        cnn = ' / '.join(
            filter(
                None,
                [
                    f"N.Art : {self.entity_info.get('nai')}" if self.entity_info.get('nai') else None,
                    f"NIS : {self.entity_info.get('nis')}" if self.entity_info.get('nis') else None,
                ],
            )
        )
        if cnn:
            c_lines.append(cnn)
        if doc_type not in (AppConstants.DOC_BV, AppConstants.DOC_BA) and self.entity_info.get('rib'):
            c_lines.append(f"RIB : {self.smart_text(self.entity_info.get('rib'))}")
        if self.doc_info.get('payment_method'):
            c_lines.append(f"Mode de Paiement : {self.smart_text(self.doc_info.get('payment_method'))}")

        ent_txt = f"{self.entity_info.get('label', 'Client')} : {self.smart_text(self.entity_info.get('name', ''))}"
        if c_lines:
            ent_txt += '\n' + '\n'.join(c_lines)
        self.set_font('arabic', 'B', 9)
        self.multi_cell(95, 5, ent_txt, border=1, align='L')
        self.set_y(max(y_inv_end, self.get_y()))
        self.ln(8)

    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', align='C')

    def _draw_final_page_footer_content(self):
        doc_type = self.doc_info.get('doc_type')
        if doc_type in AppConstants.QTY_ONLY_DOC_TYPES:
            self.ln(10)
            y_s = self.get_y()
            if y_s + 35 > self.h - self.b_margin:
                self.add_page()
                y_s = self.get_y()
            self.set_font('helvetica', 'B', 10)
            self.rect(self.l_margin, y_s, 80, 25)
            self.set_xy(self.l_margin, y_s + 2)
            self.cell(80, 5, 'Visa Expéditeur' if doc_type == AppConstants.DOC_TR else 'Signature Magasinier', align='C')
            xr = self.w - self.r_margin - 80
            self.rect(xr, y_s, 80, 25)
            self.set_xy(xr, y_s + 2)
            self.cell(80, 5, 'Visa Récepteur' if doc_type == AppConstants.DOC_TR else 'Signature Responsable', align='C')
            return

        self.ln(2)
        if self.amount_in_words:
            self.set_x(10)
            self.set_font('arabic', '', 9)
            self.multi_cell(0, 5, self.smart_text(f'Arrêté le présent document à la somme de : {self.amount_in_words}'), align='L')
            self.ln(2)

        y_cur = self.get_y()
        y_totals = y_left = y_cur
        final_label = 'Total TTC :' if doc_type in (AppConstants.DOC_FC, AppConstants.DOC_FF, AppConstants.DOC_FP, AppConstants.DOC_RC, AppConstants.DOC_RF) else 'Net à Payer :'
        rows_lbl = [
            ('Total HT :', 'total_ht'),
            ('Total TVA :', 'total_tva'),
            ('Total Remises :', 'total_discount'),
            ('Droit de Timbre :', 'stamp_duty'),
            (final_label, 'final_total'),
        ]
        if doc_type in (AppConstants.DOC_BV, AppConstants.DOC_BA):
            rows_lbl = [r for r in rows_lbl if r[1] != 'total_tva']

        start_x = self.w - self.r_margin - 72
        self.set_y(y_cur)
        for lbl_txt, key in rows_lbl:
            val = self.totals.get(key)
            if val is None:
                continue
            d_val = FinancialEngine.to_decimal(val)
            if d_val == 0 and key not in ('final_total', 'total_ht'):
                continue
            is_final = key == 'final_total'
            self.set_font('arabic', 'B' if is_final else '', 10)
            self.set_x(start_x)
            self.cell(34, 6, lbl_txt, border=1, align='C')
            self.cell(38, 6, FinancialEngine.format_money(d_val), border=1, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        ent_name = self.entity_info.get('name', '')
        excluded = (AppConstants.DOC_BI, AppConstants.DOC_DP, AppConstants.DOC_FP, AppConstants.DOC_TR, AppConstants.DOC_RC, AppConstants.DOC_RF, AppConstants.DOC_BE, AppConstants.DOC_BS)
        show_pay = doc_type not in excluded and ent_name not in (AppConstants.DEFAULT_CLIENT_NAME, AppConstants.DEFAULT_SUPPLIER_NAME)

        if self.payment_info and show_pay:
            paid_amt = FinancialEngine.to_decimal(self.payment_info.get('amount', 0))
            self.set_font('arabic', 'B', 10)
            self.set_x(start_x)
            self.cell(34, 6, 'Montant Payé :', border=1, align='C')
            self.cell(38, 6, FinancialEngine.format_money(paid_amt), border=1, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            final_tot = FinancialEngine.to_decimal(self.totals.get('final_total', 0))
            if paid_amt < final_tot:
                self.set_x(start_x)
                self.cell(34, 6, 'Reste à Payer :', border=1, align='C')
                self.cell(38, 6, FinancialEngine.format_money(final_tot - paid_amt), border=1, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        y_totals = self.get_y()

        self.set_y(y_cur)
        self.set_x(10)
        if self.balance_data and show_pay:
            self.set_font('arabic', 'B', 10)
            self.cell(60, 6, 'Situation du Solde :', border='B', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            self.set_font('arabic', '', 9)
            self.ln(2)
            self.cell(0, 5, f"Ancien Solde   : {FinancialEngine.format_money(self.balance_data['old_balance'])}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            self.cell(0, 5, f"Opération      : {FinancialEngine.format_money(self.balance_data['transaction_amount'])}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            self.set_font('arabic', 'B', 9)
            self.cell(0, 5, f"Nouveau Solde  : {FinancialEngine.format_money(self.balance_data['new_balance'])}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            y_left = self.get_y()
        self.set_y(max(y_totals, y_left) + 5)

    def draw_table_with_fill(self, headers, data, col_widths):
        full_footer_margin = 85
        simple_footer_margin = 20

        def draw_v_lines(sy, ey):
            x = self.l_margin
            self.line(x, sy, x, ey)
            for w in col_widths:
                x += w
                self.line(x, sy, x, ey)

        self.set_font('arabic', 'B', 9)
        self.set_fill_color(230, 230, 230)
        for i, h in enumerate(headers):
            self.cell(col_widths[i], 8, self.smart_text(h), border=1, align='C', fill=True)
        self.ln()
        body_start_y = self.get_y()
        self.set_font('arabic', '', 9)

        for idx, row in enumerate(data):
            is_last = idx == len(data) - 1
            row_h = 7
            des_idx = 2
            if len(row) > des_idx:
                lines = self.multi_cell(col_widths[des_idx], 7, self.smart_text(row[des_idx]), dry_run=True, output='LINES')
                row_h = max(7, len(lines) * 7)
            req_margin = full_footer_margin if is_last else simple_footer_margin
            if self.get_y() + row_h > self.h - req_margin:
                draw_v_lines(body_start_y, self.get_y())
                self.cell(sum(col_widths), 0, '', border='T')
                self.add_page()
                self.set_font('arabic', 'B', 9)
                for i, h in enumerate(headers):
                    self.cell(col_widths[i], 8, self.smart_text(h), border=1, align='C', fill=True)
                self.ln()
                self.set_font('arabic', '', 9)
                body_start_y = self.get_y()

            yb, xb = self.get_y(), self.get_x()
            for i, (cell_data, w) in enumerate(zip(row, col_widths)):
                self.set_xy(xb + sum(col_widths[:i]), yb)
                self.multi_cell(w, 7, self.smart_text(str(cell_data)), border=0, align='L' if i == des_idx else 'C')
            self.set_y(yb + row_h)

        bottom_y = self.h - full_footer_margin
        cur_y = self.get_y()
        if cur_y < bottom_y:
            draw_v_lines(body_start_y, bottom_y)
            self.set_y(bottom_y)
        else:
            draw_v_lines(body_start_y, cur_y)
        self.cell(sum(col_widths), 0, '', border='T')
        self._draw_final_page_footer_content()

    def draw_financial_receipt(self, op_label: str, op_datetime: str, note_txt: str, amount_val, recipient_sig: str = ''):
        x0 = self.l_margin
        w_box = self.w - self.l_margin - self.r_margin
        y0 = self.get_y() + 4

        self.set_draw_color(90, 110, 140)
        self.set_line_width(0.35)

        hdr_h = 14
        self.set_fill_color(235, 240, 247)
        self.rect(x0, y0, w_box, hdr_h, style='DF')

        doc_num = self.doc_info.get('doc_number', '')
        self.set_font('arabic', 'B', 11)
        self.set_xy(x0 + 4, y0 + 3)
        self.cell(105, 8, self.smart_text(f'{op_label.upper()}   N° : {doc_num}'), border=0, align='L')

        amt_w, amt_h = 70, 9
        amt_x = x0 + w_box - amt_w - 4
        amt_y = y0 + 2.5
        self.set_fill_color(255, 255, 255)
        self.rect(amt_x, amt_y, amt_w, amt_h, style='DF')
        self.set_font('arabic', 'B', 11)
        self.set_xy(amt_x, amt_y)
        self.cell(amt_w, amt_h, f'# {FinancialEngine.format_money(amount_val)} #', border=0, align='C')

        self.set_draw_color(195, 205, 220)
        self.set_line_width(0.25)
        cur_y = y0 + hdr_h + 6

        dt_code = self.doc_info.get('doc_type', '')
        ent_name = self.entity_info.get('name', '')
        if dt_code == AppConstants.DOC_SUPPLIER_PAY:
            lbl_party, lbl_w = "Payé à l'ordre de :", 34
        else:
            lbl_party, lbl_w = 'Reçu de :', 19
        self.set_font('arabic', '', 10)
        self.set_xy(x0 + 5, cur_y)
        self.cell(lbl_w, 6, lbl_party, border=0, align='L')
        self.set_font('arabic', 'B', 10.5)
        self.cell(w_box - lbl_w - 10, 6, self.smart_text(ent_name), border=0, align='L')
        cur_y += 7
        self.line(x0 + 5, cur_y, x0 + w_box - 5, cur_y)
        cur_y += 5

        self.set_font('arabic', '', 10)
        self.set_xy(x0 + 5, cur_y)
        self.cell(26, 6, 'La somme de :', border=0, align='L')
        self.set_font('arabic', 'B', 10)
        self.set_xy(x0 + 31, cur_y)
        self.multi_cell(w_box - 36, 6, self.smart_text(self.amount_in_words), border=0, align='L')
        cur_y = max(cur_y + 7, self.get_y() + 1)
        self.line(x0 + 5, cur_y, x0 + w_box - 5, cur_y)
        cur_y += 5

        clean_note = str(note_txt or '').strip() or op_label
        self.set_font('arabic', '', 10)
        self.set_xy(x0 + 5, cur_y)
        self.cell(14, 6, 'Motif :', border=0, align='L')
        self.set_xy(x0 + 19, cur_y)
        self.multi_cell(w_box - 24, 6, self.smart_text(clean_note), border=0, align='L')
        cur_y = max(cur_y + 7, self.get_y() + 1)
        self.line(x0 + 5, cur_y, x0 + w_box - 5, cur_y)
        cur_y += 5

        lieu = self.smart_text(self.store_info.get('store_address', '') or self.store_info.get('store_name', ''))
        meta_left = f'Fait à : {lieu}' if lieu else ''
        self.set_font('arabic', '', 9.5)
        self.set_xy(x0 + 5, cur_y)
        self.cell(105, 6, meta_left, border=0, align='L')
        self.set_font('arabic', 'B', 9.5)
        self.cell(w_box - 115, 6, f'Le : {op_datetime}', border=0, align='R')
        cur_y += 10

        if self.balance_data:
            strip_h = 10
            self.set_draw_color(170, 185, 205)
            self.set_fill_color(246, 248, 251)
            self.rect(x0, cur_y, w_box, strip_h, style='DF')
            col3 = w_box / 3.0
            self.set_font('arabic', '', 8.5)
            self.set_xy(x0 + 3, cur_y + 1.5)
            self.cell(col3 - 4, 7, f"Ancien Solde : {FinancialEngine.format_money(self.balance_data['old_balance'])}", border=0, align='L')
            self.cell(col3, 7, f"Opération : {FinancialEngine.format_money(self.balance_data['transaction_amount'])}", border=0, align='C')
            self.set_font('arabic', 'B', 8.5)
            self.cell(col3 - 2, 7, f"Nouveau Solde : {FinancialEngine.format_money(self.balance_data['new_balance'])}", border=0, align='R')
            cur_y += strip_h

        self.set_draw_color(60, 80, 110)
        self.set_line_width(0.55)
        self.rect(x0, y0, w_box, cur_y - y0)

        self.set_draw_color(0, 0, 0)
        self.set_line_width(0.2)


# ==============================================================================
# 6. LANGAGE KV & WIDGETS PERSONNALISÉS (SAISIE ARABE / RECYCLEVIEW FLUIDE)
# ==============================================================================
KV_BUILDER = '''
<LeftButtonsContainer>:
    adaptive_width: True
    spacing: "4dp"
    padding: "4dp"
    pos_hint: {"center_y": .5}

<RightButtonsContainer>:
    adaptive_width: True
    spacing: "8dp"
    pos_hint: {"center_y": .5}

<ProductRecycleItem>:
    orientation: 'vertical'
    size_hint_y: None
    height: dp(96)
    padding: 0
    spacing: 0

    MDCard:
        orientation: 'horizontal'
        padding: [dp(12), dp(10)]
        spacing: dp(12)
        radius: [14]
        elevation: 0
        line_color: (0.89, 0.91, 0.94, 1)
        line_width: 1
        ripple_behavior: True
        on_release: root.on_tap()
        md_bg_color: (1, 1, 1, 1)

        MDCard:
            size_hint: None, None
            size: dp(54), dp(54)
            radius: [12]
            elevation: 0
            md_bg_color: root.thumb_bg
            line_color: (0.88, 0.91, 0.95, 1)
            line_width: 1
            pos_hint: {'center_y': .5}
            ripple_behavior: True
            on_release: root.on_zoom()

            MDFloatLayout:
                MDIcon:
                    icon: root.icon_name
                    theme_text_color: "Custom"
                    text_color: root.icon_color
                    size_hint: None, None
                    size: dp(32), dp(32)
                    halign: "center"
                    valign: "middle"
                    pos_hint: {'center_x': .5, 'center_y': .5}
                    font_size: '28sp'
                    opacity: 0 if root.image_path else 1

                FitImage:
                    source: root.image_path if root.image_path else ''
                    radius: [11]
                    mipmap: True
                    opacity: 1 if root.image_path else 0
                    pos_hint: {'center_x': .5, 'center_y': .5}

        MDBoxLayout:
            orientation: 'vertical'
            size_hint_y: 1
            spacing: dp(4)

            MDLabel:
                text: root.text_name
                font_style: "Subtitle1"
                bold: True
                size_hint_y: 1
                text_size: self.size
                valign: 'center'
                halign: 'left'
                max_lines: 2
                shorten: False
                font_size: '15sp'
                theme_text_color: "Custom"
                text_color: (0.1, 0.12, 0.16, 1)
                font_name: 'ArabicFont'

            MDBoxLayout:
                orientation: 'horizontal'
                size_hint_y: None
                height: dp(28)
                spacing: dp(8)

                MDLabel:
                    text: root.text_price
                    font_style: "H6"
                    theme_text_color: "Custom"
                    text_color: root.price_color
                    bold: True
                    size_hint_x: 0.54
                    size_hint_y: 1
                    text_size: self.size
                    valign: 'center'
                    halign: 'left'
                    max_lines: 1
                    font_size: '17sp'
                    font_name: 'ArabicFont'

                MDBoxLayout:
                    size_hint_x: 0.46
                    size_hint_y: None
                    height: dp(26)
                    padding: [dp(6), 0]
                    pos_hint: {'center_y': .5}
                    canvas.before:
                        Color:
                            rgba: root.stock_bg
                        RoundedRectangle:
                            pos: self.pos
                            size: self.size
                            radius: [dp(8)]

                    MDLabel:
                        text: root.text_stock
                        theme_text_color: "Custom"
                        text_color: root.stock_color
                        size_hint_y: 1
                        text_size: self.size
                        halign: 'center'
                        valign: 'center'
                        bold: True
                        max_lines: 1
                        shorten: True
                        shorten_from: 'right'
                        font_size: '11.5sp'
                        font_name: 'ArabicFont'

<ProductRecycleView>:
    viewclass: 'ProductRecycleItem'
    RecycleBoxLayout:
        default_size: None, dp(96)
        default_size_hint: 1, None
        size_hint_y: None
        height: self.minimum_height
        orientation: 'vertical'
        spacing: dp(8)
        padding: [dp(12), dp(10), dp(12), dp(12)]

<HistoryRecycleItem>:
    orientation: "vertical"
    size_hint_y: None
    height: dp(118)
    padding: [dp(12), dp(10)]
    spacing: dp(6)
    radius: [14]
    elevation: 0
    line_color: (0.87, 0.89, 0.93, 1)
    line_width: 1
    ripple_behavior: True
    md_bg_color: root.bg_color
    on_release: root.on_tap()

    MDBoxLayout:
        orientation: "horizontal"
        size_hint_y: None
        height: dp(36)
        spacing: dp(10)

        MDFloatLayout:
            size_hint: None, None
            size: dp(36), dp(36)
            pos_hint: {"center_y": .5}
            canvas.before:
                Color:
                    rgba: (1, 1, 1, 0.92)
                RoundedRectangle:
                    pos: self.pos
                    size: self.size
                    radius: [dp(18)]

            MDIcon:
                icon: root.icon_name
                theme_text_color: "Custom"
                text_color: root.icon_color
                size_hint: None, None
                size: dp(22), dp(22)
                halign: "center"
                valign: "middle"
                pos_hint: {"center_x": .5, "center_y": .5}
                font_size: "20sp"

        MDBoxLayout:
            orientation: "vertical"
            spacing: dp(1)
            pos_hint: {"center_y": .5}

            MDLabel:
                text: root.text_type
                bold: True
                font_style: "Subtitle1"
                font_size: "15sp"
                theme_text_color: "Primary"
                font_name: 'ArabicFont'
                halign: "left"
                size_hint_y: None
                height: dp(20)

            MDLabel:
                text: root.text_ref
                font_style: "Caption"
                font_size: "11.5sp"
                theme_text_color: "Secondary"
                font_name: 'ArabicFont'
                halign: "left"
                size_hint_y: None
                height: dp(15) if root.text_ref else 0
                opacity: 1 if root.text_ref else 0

    MDLabel:
        text: root.text_entity
        font_style: "Subtitle1"
        font_size: "15sp"
        bold: True
        theme_text_color: "Custom"
        text_color: (0.1, 0.12, 0.16, 1)
        font_name: 'ArabicFont'
        size_hint_y: 1
        text_size: self.size
        valign: "center"
        halign: "left"
        max_lines: 2
        shorten: False

    MDBoxLayout:
        orientation: "horizontal"
        size_hint_y: None
        height: dp(22) if (root.text_date or root.text_amount) else 0
        opacity: 1 if (root.text_date or root.text_amount) else 0

        MDLabel:
            text: root.text_date
            font_style: "Caption"
            font_size: "11.5sp"
            theme_text_color: "Hint"
            halign: "left"
            valign: "center"
            size_hint_x: 0.45
            font_name: 'ArabicFont'

        MDLabel:
            text: root.text_amount
            font_style: "Subtitle1"
            bold: True
            theme_text_color: "Custom"
            text_color: root.icon_color
            halign: "right"
            valign: "center"
            font_size: "16sp"
            size_hint_x: 0.55
            font_name: 'ArabicFont'

<HistoryRecycleView>:
    viewclass: 'HistoryRecycleItem'
    RecycleBoxLayout:
        default_size: None, dp(118)
        default_size_hint: 1, None
        size_hint_y: None
        height: self.minimum_height
        orientation: 'vertical'
        spacing: dp(8)
        padding: [dp(2), dp(4)]

<EntityRecycleItem>:
    orientation: "horizontal"
    size_hint_y: None
    height: dp(86)
    padding: [dp(12), dp(8)]
    spacing: dp(12)
    ripple_behavior: True
    md_bg_color: root.bg_color
    line_color: (0.89, 0.91, 0.94, 1)
    line_width: 1
    radius: [12]
    elevation: 0
    on_release: root.on_tap()

    MDFloatLayout:
        size_hint: None, None
        size: dp(44), dp(44)
        pos_hint: {"center_y": .5}
        canvas.before:
            Color:
                rgba: root.icon_bg
            RoundedRectangle:
                pos: self.pos
                size: self.size
                radius: [dp(22)]

        MDIcon:
            icon: root.icon_name
            theme_text_color: "Custom"
            text_color: root.icon_color
            size_hint: None, None
            size: dp(26), dp(26)
            halign: "center"
            valign: "middle"
            pos_hint: {"center_x": .5, "center_y": .5}
            font_size: "24sp"

    MDBoxLayout:
        orientation: "vertical"
        size_hint_x: 1
        size_hint_y: 1
        spacing: dp(2)

        MDLabel:
            text: root.text_name
            bold: True
            font_style: "Subtitle1"
            font_size: "15sp"
            font_name: 'ArabicFont'
            theme_text_color: "Custom"
            text_color: (0.1, 0.12, 0.16, 1)
            size_hint_y: 1
            text_size: self.size
            valign: "center"
            halign: "left"
            max_lines: 2
            shorten: False

        MDLabel:
            text: root.text_balance
            font_style: "Caption"
            font_name: 'ArabicFont'
            markup: True
            theme_text_color: "Secondary"
            size_hint_y: None
            height: dp(18)

    MDIcon:
        icon: "chevron-right"
        theme_text_color: "Hint"
        font_size: "22sp"
        size_hint_x: None
        width: dp(22)
        pos_hint: {"center_y": .5}

<EntityRecycleView>:
    viewclass: 'EntityRecycleItem'
    RecycleBoxLayout:
        default_size: None, dp(86)
        default_size_hint: 1, None
        size_hint_y: None
        height: self.minimum_height
        orientation: 'vertical'
        spacing: dp(6)
        padding: [dp(2), dp(4)]

<MgmtEntityRecycleItem>:
    orientation: "horizontal"
    size_hint_y: None
    height: dp(88)
    padding: [dp(12), dp(8)]
    spacing: dp(10)
    ripple_behavior: True
    md_bg_color: root.bg_color
    line_color: (0.89, 0.91, 0.94, 1)
    line_width: 1
    radius: [12]
    elevation: 0
    on_release: root.on_pay()

    MDFloatLayout:
        size_hint: None, None
        size: dp(42), dp(42)
        pos_hint: {"center_y": .5}
        canvas.before:
            Color:
                rgba: (0.93, 0.95, 0.98, 1)
            RoundedRectangle:
                pos: self.pos
                size: self.size
                radius: [dp(21)]

        MDIcon:
            icon: "account-circle"
            theme_text_color: "Custom"
            text_color: (0.35, 0.42, 0.52, 1)
            size_hint: None, None
            size: dp(26), dp(26)
            halign: "center"
            valign: "middle"
            pos_hint: {"center_x": .5, "center_y": .5}
            font_size: "24sp"

    MDBoxLayout:
        orientation: "vertical"
        size_hint_x: 1
        size_hint_y: 1
        spacing: dp(2)

        MDLabel:
            text: root.text_name
            bold: True
            font_style: "Subtitle1"
            font_size: "15sp"
            font_name: 'ArabicFont'
            theme_text_color: "Custom"
            text_color: (0.1, 0.12, 0.16, 1)
            size_hint_y: 1
            text_size: self.size
            valign: "center"
            halign: "left"
            max_lines: 2
            shorten: False

        MDLabel:
            text: root.text_balance
            font_style: "Caption"
            font_name: 'ArabicFont'
            markup: True
            theme_text_color: "Secondary"
            halign: "left"
            size_hint_y: None
            height: dp(18)

    MDCard:
        id: btn_hist
        size_hint: None, None
        size: dp(38), dp(38)
        radius: [19]
        elevation: 0
        md_bg_color: (0.88, 0.96, 0.96, 1)
        ripple_behavior: True
        pos_hint: {"center_y": .5}
        on_release: root.on_history()

        MDFloatLayout:
            MDIcon:
                icon: "history"
                theme_text_color: "Custom"
                text_color: (0, 0.5, 0.55, 1)
                size_hint: None, None
                size: dp(22), dp(22)
                halign: "center"
                valign: "middle"
                pos_hint: {"center_x": .5, "center_y": .5}
                font_size: "21sp"

<MgmtEntityRecycleView>:
    viewclass: 'MgmtEntityRecycleItem'
    RecycleBoxLayout:
        default_size: None, dp(88)
        default_size_hint: 1, None
        size_hint_y: None
        height: self.minimum_height
        orientation: 'vertical'
        spacing: dp(6)
        padding: [dp(2), dp(4)]

<CartRecycleItem>:
    orientation: "horizontal"
    size_hint_y: None
    height: dp(96)
    padding: [dp(12), dp(10), dp(10), dp(10)]
    spacing: dp(10)
    md_bg_color: 1, 1, 1, 1
    line_color: (0.89, 0.91, 0.94, 1)
    line_width: 1
    radius: [14]
    elevation: 0
    ripple_behavior: True
    on_release: root.on_tap()

    MDFloatLayout:
        size_hint: None, None
        size: dp(42), dp(42)
        pos_hint: {"center_y": .5}
        canvas.before:
            Color:
                rgba: (0.92, 0.95, 1, 1)
            RoundedRectangle:
                pos: self.pos
                size: self.size
                radius: [dp(12)]

        MDIcon:
            icon: root.badge_icon
            theme_text_color: "Custom"
            text_color: (0.1, 0.45, 0.85, 1)
            size_hint: None, None
            size: dp(24), dp(24)
            halign: "center"
            valign: "middle"
            pos_hint: {"center_x": .5, "center_y": .5}
            font_size: "22sp"

    MDBoxLayout:
        orientation: "vertical"
        size_hint_y: 1
        spacing: dp(4)

        MDLabel:
            text: root.text_name
            font_style: "Subtitle1"
            font_size: "15sp"
            bold: True
            theme_text_color: "Custom"
            text_color: (0.1, 0.12, 0.16, 1)
            size_hint_y: 1
            text_size: self.size
            valign: "center"
            halign: "left"
            max_lines: 2
            shorten: False
            font_name: 'ArabicFont'

        MDBoxLayout:
            orientation: "horizontal"
            size_hint_y: None
            height: dp(24)
            spacing: dp(6)

            MDLabel:
                text: root.text_details
                font_size: "13sp"
                theme_text_color: "Custom"
                text_color: (0.42, 0.46, 0.52, 1) if root.text_subtotal else root.details_color
                bold: False if root.text_subtotal else True
                size_hint_x: 0.56 if root.text_subtotal else 1
                size_hint_y: 1
                text_size: self.size
                valign: "center"
                halign: "left"
                max_lines: 1
                shorten: True
                shorten_from: "right"
                font_name: 'ArabicFont'

            MDLabel:
                text: root.text_subtotal
                font_size: "14sp"
                theme_text_color: "Custom"
                text_color: root.details_color
                bold: True
                size_hint_x: 0.44 if root.text_subtotal else 0
                size_hint_y: 1
                text_size: self.size
                valign: "center"
                halign: "right"
                max_lines: 1
                opacity: 1 if root.text_subtotal else 0
                font_name: 'ArabicFont'

    MDFloatLayout:
        size_hint: None, None
        size: dp(38) if root.can_delete else 0, dp(38) if root.can_delete else 0
        opacity: 1 if root.can_delete else 0
        disabled: not root.can_delete
        pos_hint: {"center_y": .5}
        canvas.before:
            Color:
                rgba: (1, 0.93, 0.93, 1) if root.can_delete else (0, 0, 0, 0)
            RoundedRectangle:
                pos: self.pos
                size: self.size
                radius: [dp(10)]

        MDIconButton:
            icon: "trash-can-outline"
            theme_text_color: "Custom"
            text_color: (0.85, 0.14, 0.14, 1)
            pos_hint: {"center_x": .5, "center_y": .5}
            icon_size: "20sp"
            disabled: not root.can_delete
            on_release: root.on_delete()

<CartRecycleView>:
    viewclass: 'CartRecycleItem'
    RecycleBoxLayout:
        default_size: None, dp(96)
        default_size_hint: 1, None
        size_hint_y: None
        height: self.minimum_height
        orientation: 'vertical'
        spacing: dp(8)
        padding: [dp(12), dp(10), dp(12), dp(12)]
'''


class NoMenuTextField(MDTextField):
    def _show_cut_copy_paste(self, pos, selection, mode=None):
        pass

    def on_double_tap(self):
        pass


class SmartTextField(MDTextField):
    """Champ de texte conservant le texte logique brut tout en affichant l'arabe mis en forme."""

    def __init__(self, **kwargs):
        self._raw_text = ''
        self.base_direction = 'ltr'
        self.halign = 'left'
        super().__init__(**kwargs)
        self.font_name = 'ArabicFont'
        self.font_name_hint_text = 'ArabicFont'
        self.keyboard_suggestions = False
        if self.text:
            self._raw_text = self.text
            self._update_display()

    def _apply_input_filter(self, substring: str, from_undo: bool = False) -> str:
        flt = getattr(self, 'input_filter', None)
        if not flt or not substring:
            return substring
        if callable(flt):
            return flt(substring, from_undo)
        if flt == 'int':
            allowed = '0123456789' if self._raw_text else '-0123456789'
            return ''.join(ch for ch in substring if ch in allowed)
        if flt == 'float':
            out = []
            has_dot = '.' in self._raw_text
            for idx, ch in enumerate(substring):
                if ch in '0123456789':
                    out.append(ch)
                elif ch == '.' and not has_dot:
                    out.append(ch)
                    has_dot = True
                elif ch == '-' and not self._raw_text and idx == 0 and not out:
                    out.append(ch)
            return ''.join(out)
        return substring

    def set_raw_value(self, value: str):
        raw = str(value or '')
        if getattr(self, 'input_filter', None):
            prev = self._raw_text
            self._raw_text = ''
            raw = self._apply_input_filter(raw, False)
            self._raw_text = prev
        self._raw_text = raw
        self._update_display()

    def insert_text(self, substring, from_undo=False):
        filtered = self._apply_input_filter(substring, from_undo)
        if not filtered:
            return
        self._raw_text += filtered
        self._update_display()

    def do_backspace(self, from_undo=False, mode='bkspc'):
        if not self._raw_text:
            return
        self._raw_text = self._raw_text[:-1]
        self._update_display()

    def _update_display(self):
        self.text = ar_visual(self._raw_text, _INPUT_RESHAPER) if self._raw_text else ''
        self._update_alignment(self._raw_text)
        Clock.schedule_once(self._set_cursor_to_end, 0)

    def _set_cursor_to_end(self, dt):
        self.cursor = (len(self.text), 0)

    def _update_alignment(self, text: str):
        if not text:
            self.halign = 'left'
            self.base_direction = 'ltr'
            return
        has_arabic = bool(_AR_RAW_RE.search(text))
        self.halign = 'right' if has_arabic else 'left'
        self.base_direction = 'rtl' if has_arabic else 'ltr'

    def get_value(self) -> str:
        if not self._raw_text and self.text:
            return self.text
        return self._raw_text

    def on_text(self, instance, value):
        if value == '' and self._raw_text != '':
            self._raw_text = ''


class LeftButtonsContainer(ILeftBody, MDBoxLayout):
    adaptive_width = True


class RightButtonsContainer(IRightBodyTouch, MDBoxLayout):
    adaptive_width = True


class CartRecycleItem(RecycleDataViewBehavior, MDCard):
    text_name = StringProperty('')
    text_details = StringProperty('')
    text_subtotal = StringProperty('')
    details_color = ListProperty([0.4, 0.4, 0.4, 1])
    badge_icon = StringProperty('package-variant')
    can_delete = BooleanProperty(True)
    raw_data = ObjectProperty(None, allownone=True)

    def refresh_view_attrs(self, rv, index, data):
        self.index = index
        self.text_name = ar_visual(data.get('name', ''))
        self.text_details = data.get('details', '')
        self.text_subtotal = data.get('subtotal', '')
        self.details_color = data.get('d_color', [0.4, 0.4, 0.4, 1])
        self.badge_icon = data.get('badge_icon', 'package-variant')
        self.can_delete = bool(data.get('can_delete', True))
        self.raw_data = data.get('raw_item')
        return super().refresh_view_attrs(rv, index, data)

    def on_tap(self):
        app = MDApp.get_running_app()
        if app and self.raw_data:
            app.edit_cart_item(self.raw_data)

    def on_delete(self):
        if not self.can_delete:
            return
        app = MDApp.get_running_app()
        if app and self.raw_data:
            app.remove_from_cart(self.raw_data)


class ProductRecycleItem(RecycleDataViewBehavior, MDBoxLayout):
    index = None
    text_name = StringProperty('')
    text_price = StringProperty('')
    text_stock = StringProperty('')
    icon_name = StringProperty('package-variant')
    icon_color = ListProperty([0, 0, 0, 1])
    price_color = ListProperty([0, 0, 0, 1])
    stock_color = ListProperty([0.15, 0.2, 0.25, 1])
    stock_bg = ListProperty([0.94, 0.95, 0.97, 1])
    thumb_bg = ListProperty([0.94, 0.96, 1, 1])
    image_path = StringProperty('')
    product_data = ObjectProperty(None, allownone=True)

    def refresh_view_attrs(self, rv, index, data):
        self.index = index
        self.text_name = data.get('text_name', '')
        self.text_price = data.get('text_price', '')
        self.text_stock = data.get('text_stock', '')
        self.icon_name = data.get('icon_name', 'package-variant')
        self.icon_color = data.get('icon_color', [0, 0, 0, 1])
        self.price_color = data.get('price_color', [0, 0, 0, 1])
        self.stock_color = data.get('stock_color', [0.15, 0.2, 0.25, 1])
        self.stock_bg = data.get('stock_bg', [0.94, 0.95, 0.97, 1])
        self.thumb_bg = data.get('thumb_bg', [0.94, 0.96, 1, 1])
        self.image_path = data.get('image_path', '')
        self.product_data = data.get('raw_data')
        return super().refresh_view_attrs(rv, index, data)

    def on_tap(self):
        app = MDApp.get_running_app()
        if app and self.product_data:
            app.open_add_to_cart_dialog(self.product_data, app.current_mode)

    def on_zoom(self):
        if self.image_path:
            app = MDApp.get_running_app()
            if app:
                app.show_zoomed_image(self.image_path, self.text_name)
        else:
            self.on_tap()


class HistoryRecycleItem(RecycleDataViewBehavior, MDCard):
    index = None
    text_type = StringProperty('')
    text_ref = StringProperty('')
    text_entity = StringProperty('')
    text_date = StringProperty('')
    text_amount = StringProperty('')
    icon_name = StringProperty('file')
    icon_color = ColorProperty([0, 0, 0, 1])
    bg_color = ColorProperty([1, 1, 1, 1])
    op_id = ObjectProperty(None, allownone=True)
    is_transfer = BooleanProperty(False)

    def refresh_view_attrs(self, rv, index, data):
        self.index = index
        self.text_type = ar_visual(data.get('type_str', ''))
        self.text_ref = ar_visual(data.get('ref_str', ''))
        self.text_entity = ar_visual(data.get('entity_str', ''))
        self.text_date = data.get('date_str', '')
        self.text_amount = data.get('amount_text', '')
        self.icon_name = data.get('icon', 'file')
        self.icon_color = data.get('icon_color', [0, 0, 0, 1])
        self.bg_color = data.get('bg_color', [1, 1, 1, 1])
        self.op_id = data.get('op_id')
        self.is_transfer = bool(data.get('is_transfer', False))
        return super().refresh_view_attrs(rv, index, data)

    def on_tap(self):
        app = MDApp.get_running_app()
        if app and self.op_id is not None:
            app.view_transaction_details(self.op_id, self.is_transfer)


class MgmtEntityRecycleItem(RecycleDataViewBehavior, MDCard):
    index = None
    text_name = StringProperty('')
    text_balance = StringProperty('')
    bg_color = ColorProperty([1, 1, 1, 1])
    entity_data = ObjectProperty(None, allownone=True)
    _long_press_event = None
    _is_long_press = False
    _start_pos = (0, 0)

    def refresh_view_attrs(self, rv, index, data):
        self.index = index
        self.text_name = ar_visual(data.get('raw_name', ''))
        self.text_balance = data.get('balance_text', '')
        self.bg_color = data.get('bg_color', [1, 1, 1, 1])
        self.entity_data = data.get('raw_data')
        return super().refresh_view_attrs(rv, index, data)

    def _cancel_long_press(self):
        if self._long_press_event:
            self._long_press_event.cancel()
            self._long_press_event = None

    def on_touch_down(self, touch):
        if self.collide_point(*touch.pos):
            self._is_long_press = False
            self._cancel_long_press()
            btn_hist = self.ids.get('btn_hist')
            if not (btn_hist and btn_hist.collide_point(*touch.pos)):
                self._start_pos = touch.pos
                self._long_press_event = Clock.schedule_once(lambda dt: self._trigger_long_press(), 0.55)
        return super().on_touch_down(touch)

    def on_touch_move(self, touch):
        if self._long_press_event:
            if abs(touch.x - self._start_pos[0]) > dp(20) or abs(touch.y - self._start_pos[1]) > dp(20):
                self._cancel_long_press()
        return super().on_touch_move(touch)

    def on_touch_up(self, touch):
        self._cancel_long_press()
        if self._is_long_press:
            self._is_long_press = False
            self.state = 'normal'
            touch.ungrab(self)
            return True
        return super().on_touch_up(touch)

    def _trigger_long_press(self):
        self._long_press_event = None
        app = MDApp.get_running_app()
        if app and getattr(app, 'entity_hist_dialog', None) and app.entity_hist_dialog.parent:
            return
        self._is_long_press = True
        self.state = 'normal'
        if app and self.entity_data:
            app.open_entity_edit_menu(self.entity_data)

    def on_pay(self):
        self._cancel_long_press()
        if self._is_long_press:
            return
        app = MDApp.get_running_app()
        if app and getattr(app, 'options_dialog', None) and app.options_dialog.parent:
            return
        if app and self.entity_data:
            app.start_direct_payment_from_manager(self.entity_data)

    def on_history(self):
        self._cancel_long_press()
        self._is_long_press = False
        self.state = 'normal'
        app = MDApp.get_running_app()
        if app and self.entity_data:
            app.open_entity_history_dialog(self.entity_data)


class EntityRecycleItem(RecycleDataViewBehavior, MDCard):
    index = None
    text_name = StringProperty('')
    text_balance = StringProperty('')
    icon_name = StringProperty('account')
    icon_color = ListProperty([0, 0, 0, 1])
    bg_color = ColorProperty([1, 1, 1, 1])
    icon_bg = ColorProperty([0.94, 0.96, 1, 1])
    entity_data = ObjectProperty(None, allownone=True)

    def refresh_view_attrs(self, rv, index, data):
        self.index = index
        self.text_name = ar_visual(data.get('raw_name', ''))
        self.text_balance = data.get('balance_text', '')
        self.icon_name = data.get('icon', 'account')
        self.icon_color = data.get('icon_color', [0, 0, 0, 1])
        self.bg_color = data.get('bg_color', [1, 1, 1, 1])
        self.icon_bg = data.get('icon_bg', [0.94, 0.96, 1, 1])
        self.entity_data = data.get('raw_data')
        return super().refresh_view_attrs(rv, index, data)

    def on_tap(self):
        app = MDApp.get_running_app()
        if app and self.entity_data:
            app.select_entity_from_rv(self.entity_data)


class _SmoothRecycleView(RecycleView):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.data = []
        self.loading_lock = False
        self._restoring_scroll = False
        self._scroll_gen = 0
        self.do_scroll_x = False
        self.do_scroll_y = True
        self.scroll_type = ['bars', 'content']
        self.bar_width = dp(4)
        self.scroll_distance = dp(10)
        self.scroll_timeout = 150
        Clock.schedule_once(self._configure_scroll_effect, 0)

    def _configure_scroll_effect(self, dt=0):
        if hasattr(self, 'effect_y') and self.effect_y:
            self.effect_y.friction = 0.18
            if hasattr(self.effect_y, 'min_velocity'):
                self.effect_y.min_velocity = dp(12)

    def on_scroll_move(self, touch):
        res = super().on_scroll_move(touch)
        self._clamp_velocity()
        return res

    def on_touch_up(self, touch):
        res = super().on_touch_up(touch)
        self._clamp_velocity()
        return res

    def _clamp_velocity(self):
        if hasattr(self, 'effect_y') and self.effect_y:
            max_v = dp(1800)
            vel = getattr(self.effect_y, 'velocity', 0)
            if vel > max_v:
                self.effect_y.velocity = max_v
            elif vel < -max_v:
                self.effect_y.velocity = -max_v


class CartRecycleView(_SmoothRecycleView):
    pass


class ProductRecycleView(_SmoothRecycleView):
    def on_scroll_y(self, instance, value):
        if getattr(self, '_restoring_scroll', False):
            return
        app = MDApp.get_running_app()
        if app and value <= 0.15 and not self.loading_lock and not app.is_loading_more and getattr(app, 'products_has_more', True) and len(self.data) >= 15:
            self.loading_lock = True
            app.load_more_products(reset=False)


class HistoryRecycleView(_SmoothRecycleView):
    def on_scroll_y(self, instance, value):
        if getattr(self, '_restoring_scroll', False):
            return
        app = MDApp.get_running_app()
        if app and value <= 0.12 and not self.loading_lock and not app.is_loading_history and getattr(app, 'history_has_more', True) and len(self.data) >= 15:
            self.loading_lock = True
            app.load_more_history()


class EntityRecycleView(_SmoothRecycleView):
    def on_scroll_y(self, instance, value):
        if getattr(self, '_restoring_scroll', False):
            return
        app = MDApp.get_running_app()
        if app and value <= 0.12 and not self.loading_lock and not app.is_loading_entities and getattr(app, 'entities_has_more', True) and len(self.data) >= 15:
            self.loading_lock = True
            app.load_more_entities(reset=False)


class MgmtEntityRecycleView(_SmoothRecycleView):
    def on_scroll_y(self, instance, value):
        if getattr(self, '_restoring_scroll', False):
            return
        app = MDApp.get_running_app()
        if app and value <= 0.12 and not self.loading_lock and not app.is_loading_entities and getattr(app, 'entities_has_more', True) and len(self.data) >= 15:
            self.loading_lock = True
            app.load_more_entities(reset=False)


# ==============================================================================
# 7. APPLICATION PRINCIPALE (StockApp - MAGPRO MOBILE v8.1.4)
# ==============================================================================
class StockApp(MDApp):
    is_seller_mode = BooleanProperty(False)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.cart = []
        self.current_mode = 'sale'
        self.current_user_name = AppConstants.ADMIN_USERNAME
        self.selected_location = AppConstants.LOC_STORE
        self.selected_entity = None
        self.selected_family_filter = AppConstants.FAM_ALL

        self.editing_tx_id = None
        self.editing_doc_type = None
        self.editing_payment_amount = Decimal('0.00')
        self.editing_payment_method = AppConstants.PAY_DEFAULT
        self.editing_global_discount = Decimal('0.00')
        self.current_editing_date = None
        self.is_transaction_in_progress = False

        self.current_page_offset = 0
        self.batch_size = 50
        self.is_loading_more = False
        self.products_has_more = True

        self.history_page_offset = 0
        self.history_batch_size = 50
        self.is_loading_history = False
        self.history_has_more = True
        self.history_start_date = datetime.now().date()
        self.history_end_date = datetime.now().date()
        self.current_history_filter = AppConstants.FAM_ALL

        self.entity_page_offset = 0
        self.is_loading_entities = False
        self.entities_has_more = True
        self.current_entity_type_mgmt = AppConstants.ENTITY_CLIENT
        self.current_entity_sort = 'name'
        self.active_entity_rv = None

        self._search_event = None
        self._entity_search_event = None
        self._notify_event = None
        self._last_click_time = 0
        self._stats_seq = 0
        self._history_fetch_seq = 0
        self._entity_fetch_seq = 0
        self.is_showing_alert = False
        self.tone_gen = None
        self.db = None

        self.dialog = None
        self.rv_products = None
        self.status_bar_label = None
        self.status_bar_icon = None
        self.status_bar_bg = None

    def fix_text(self, text) -> str:
        return ar_visual(text)



    # --------------------------------------------------------------------------
    # 7.0 FABRIQUE UI UNIFIÉE (DESIGN SYSTEM 10/10 SANS DUPLICATION)
    # --------------------------------------------------------------------------
    def _fit_dialog(self, dlg, top_pad: float = 16, side_pad: float = 14, bottom_pad: float = 12, spacer_top: float = 8):
        if not dlg:
            return
        try:
            if 'container' in dlg.ids:
                dlg.ids.container.padding = [dp(side_pad), dp(top_pad), dp(side_pad), dp(bottom_pad)]
            if 'spacer_top_box' in dlg.ids:
                dlg.ids.spacer_top_box.padding = [0, dp(spacer_top), 0, 0]
        except Exception:
            pass

    def _ui_card_btn(
        self,
        text: str,
        bg_color,
        text_color=(1, 1, 1, 1),
        icon: str | None = None,
        height: float = 48,
        radius: float = 12,
        size_hint_x: float = 1,
        font_size: str = '14.5sp',
        border_color=None,
        elevation: int = 0,
        on_release=None,
    ) -> MDCard:
        btn = MDCard(
            orientation='horizontal',
            size_hint_x=size_hint_x,
            size_hint_y=None,
            height=dp(height),
            radius=[radius],
            elevation=elevation,
            md_bg_color=bg_color,
            ripple_behavior=True,
            spacing=dp(5),
            padding=[dp(8), 0],
        )
        if border_color is not None:
            btn.line_color = border_color
            btn.line_width = 1
        if on_release is not None:
            btn.bind(on_release=on_release)

        btn.add_widget(MDBoxLayout())
        ico = None
        if icon:
            ico = MDIcon(
                icon=icon,
                theme_text_color='Custom',
                text_color=text_color,
                font_size='18sp',
                size_hint=(None, None),
                size=(dp(20), dp(20)),
                halign='center',
                valign='middle',
                pos_hint={'center_y': 0.5},
            )
            btn._ico = ico
            btn.add_widget(ico)
        lbl = MDLabel(
            text=text,
            bold=True,
            font_size=font_size,
            theme_text_color='Custom',
            text_color=text_color,
            size_hint=(None, None),
            max_lines=1,
            shorten_from='right',
            halign='center',
            valign='middle',
            pos_hint={'center_y': 0.5},
        )
        btn._lbl = lbl
        btn.add_widget(lbl)
        btn.add_widget(MDBoxLayout())

        def _sync_btn_fit(*_):
            if btn.width <= 0:
                return
            pad_x = (btn.padding[0] + btn.padding[2]) if len(btn.padding) >= 4 else (btn.padding[0] * 2)
            ico_w = (ico.width + btn.spacing) if ico is not None else 0
            max_lbl_w = max(dp(10), btn.width - pad_x - ico_w)
            lbl.text_size = (None, None)
            lbl.texture_update()
            nat_w, nat_h = lbl.texture_size
            lbl.height = max(dp(20), nat_h)
            if nat_w > max_lbl_w:
                lbl.width = max_lbl_w
                lbl.text_size = (max_lbl_w, None)
                lbl.shorten = True
            else:
                lbl.width = nat_w
                lbl.text_size = (None, None)
                lbl.shorten = False

        btn.bind(width=_sync_btn_fit)
        lbl.bind(text=_sync_btn_fit)
        Clock.schedule_once(_sync_btn_fit, 0)
        return btn

    def _ui_icon_pill(
        self,
        icon_name: str,
        bg_color,
        icon_color,
        on_release=None,
        size_dp: float = 40,
        radius: float = 10,
        icon_size: str = '20sp',
        border_color=None,
        clickable: bool = False,
    ):
        sz = max(dp(18), dp(size_dp) * 0.55)
        is_interactive = bool(on_release is not None or clickable)

        if not is_interactive:
            from kivy.graphics import Color, Line, RoundedRectangle

            badge = MDFloatLayout(
                size_hint=(None, None),
                size=(dp(size_dp), dp(size_dp)),
                pos_hint={'center_y': 0.5},
            )
            with badge.canvas.before:
                bg_instr = Color(*bg_color)
                bg_rect = RoundedRectangle(pos=badge.pos, size=badge.size, radius=[dp(radius)])
                border_instr = None
                border_line = None
                if border_color is not None:
                    border_instr = Color(*border_color)
                    border_line = Line(
                        rounded_rectangle=(badge.x, badge.y, badge.width, badge.height, dp(radius)),
                        width=1,
                    )

            def _sync_badge(inst, val):
                bg_rect.pos = inst.pos
                bg_rect.size = inst.size
                if border_line is not None:
                    border_line.rounded_rectangle = (inst.x, inst.y, inst.width, inst.height, dp(radius))

            badge.bind(pos=_sync_badge, size=_sync_badge)
            ico = MDIcon(
                icon=icon_name,
                theme_text_color='Custom',
                text_color=icon_color,
                font_size=icon_size,
                size_hint=(None, None),
                size=(sz, sz),
                halign='center',
                valign='middle',
                pos_hint={'center_x': 0.5, 'center_y': 0.5},
            )
            badge._ico = ico
            badge._bg_color_instr = bg_instr
            badge._border_color_instr = border_instr
            badge.add_widget(ico)
            return badge

        btn = MDCard(
            size_hint=(None, None),
            size=(dp(size_dp), dp(size_dp)),
            radius=[radius],
            elevation=0,
            md_bg_color=bg_color,
            ripple_behavior=True,
            pos_hint={'center_y': 0.5},
        )
        if border_color is not None:
            btn.line_color = border_color
            btn.line_width = 1
        if on_release is not None:
            btn.bind(on_release=on_release)
        fl = MDFloatLayout()
        ico = MDIcon(
            icon=icon_name,
            theme_text_color='Custom',
            text_color=icon_color,
            font_size=icon_size,
            size_hint=(None, None),
            size=(sz, sz),
            halign='center',
            valign='middle',
            pos_hint={'center_x': 0.5, 'center_y': 0.5},
        )
        fl.add_widget(ico)
        btn._ico = ico
        btn.add_widget(fl)
        return btn

    def _ui_section_card(
        self,
        bg_color,
        border_color,
        icon_name: str,
        title_text: str,
        icon_color=None,
        right_widget=None,
    ) -> MDCard:
        card = MDCard(
            orientation='vertical',
            radius=[14],
            padding=[dp(12), dp(10)],
            spacing=dp(8),
            elevation=0,
            adaptive_height=True,
            md_bg_color=bg_color,
            line_color=border_color,
            line_width=1,
        )
        ic_c = icon_color or self.theme_cls.primary_color
        hdr = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint_y=None, height=dp(28))
        hdr.add_widget(
            MDIcon(
                icon=icon_name,
                theme_text_color='Custom',
                text_color=ic_c,
                font_size='19sp',
                size_hint=(None, None),
                size=(dp(22), dp(22)),
                pos_hint={'center_y': 0.5},
            )
        )
        hdr.add_widget(
            MDLabel(
                text=title_text,
                bold=True,
                theme_text_color='Custom',
                text_color=ic_c,
                font_style='Subtitle2',
                pos_hint={'center_y': 0.5},
            )
        )
        if right_widget is not None:
            hdr.add_widget(right_widget)
        card.add_widget(hdr)
        card.add_widget(MDSeparator(height=dp(1), color=(0, 0, 0, 0.08)))
        return card

    def _ui_wrap_label(
        self,
        text: str,
        font_style: str = 'Body2',
        halign: str = 'left',
        bold: bool = False,
        theme_color: str = 'Secondary',
        text_color=None,
        min_h: float = 20,
        font_name: str = 'ArabicFont',
    ) -> MDLabel:
        kw = {
            'text': text,
            'font_style': font_style,
            'halign': halign,
            'bold': bold,
            'theme_text_color': theme_color,
            'size_hint_x': 1,
            'size_hint_y': None,
            'height': dp(min_h),
            'font_name': font_name,
        }
        if text_color is not None:
            kw['text_color'] = text_color
        lbl = MDLabel(**kw)
        lbl.bind(
            width=lambda inst, val: setattr(inst, 'text_size', (max(1, val), None)),
            texture_size=lambda inst, val: setattr(inst, 'height', max(dp(min_h), val[1])),
        )
        return lbl

    def _ui_period_tabs(self, on_filter, on_calendar):
        class _PeriodPill(MDCard):
            text = StringProperty('')
            text_color = ColorProperty([0.35, 0.38, 0.45, 1])

            def __init__(self, **kwargs):
                super().__init__(**kwargs)
                self.size_hint_x = 1
                self.size_hint_y = 1
                self.radius = [9]
                self.elevation = 0
                self.ripple_behavior = True
                self._lbl = MDLabel(
                    text=self.text,
                    bold=True,
                    font_size='11.5sp',
                    halign='center',
                    valign='middle',
                    theme_text_color='Custom',
                    text_color=self.text_color,
                )
                self.add_widget(self._lbl)
                self.bind(
                    text=lambda inst, val: setattr(self._lbl, 'text', val),
                    text_color=lambda inst, val: setattr(self._lbl, 'text_color', val),
                )

        tabs_container = MDCard(
            orientation='horizontal',
            size_hint=(1, None),
            height=dp(42),
            radius=[12],
            elevation=0,
            md_bg_color=(0.94, 0.95, 0.97, 1),
            line_color=(0.88, 0.90, 0.94, 1),
            line_width=1,
            padding=dp(4),
            spacing=dp(4),
        )
        btns = {
            'auj': _PeriodPill(text='AUJ', on_release=lambda x: on_filter('auj')),
            'hier': _PeriodPill(text='HIE', on_release=lambda x: on_filter('hier')),
            'semaine': _PeriodPill(text='SEM', on_release=lambda x: on_filter('semaine')),
            'mois': _PeriodPill(text='MOI', on_release=lambda x: on_filter('mois')),
            'custom': _PeriodPill(text='CAL', on_release=on_calendar),
        }
        for key in ('auj', 'hier', 'semaine', 'mois', 'custom'):
            tabs_container.add_widget(btns[key])
        return tabs_container, btns

    def _activate_period_tab(self, btns: dict, active_key: str, cal_label: str = 'CAL'):
        inact_bg, act_bg = (0.94, 0.95, 0.97, 1), (0.09, 0.44, 0.84, 1)
        inact_txt, act_txt = (0.35, 0.38, 0.45, 1), (1, 1, 1, 1)
        for k, b in btns.items():
            is_act = k == active_key
            b.md_bg_color = act_bg if is_act else inact_bg
            b.text_color = act_txt if is_act else inact_txt
        if 'custom' in btns:
            btns['custom'].text = cal_label

    def _ui_filter_menu(self, title: str, options: list, current_code: str, on_select, on_delete_item=None):
        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(10),
            size_hint_y=None,
            height=min(dp(440), Window.height * 0.72),
            padding=[dp(2), dp(4), dp(2), dp(2)],
        )
        scroll = MDScrollView(size_hint_y=1, bar_width=dp(3))
        lst = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(8), padding=[dp(2), dp(4), dp(2), dp(4)])

        dlg_ref = {'dlg': None}

        for opt in options:
            label = opt['label']
            code = opt['code']
            ico_name = opt.get('icon', 'tag-outline')
            ico_col = opt.get('color', (0.45, 0.48, 0.55, 1))
            is_sel = code == current_code

            card = MDCard(
                orientation='horizontal',
                size_hint=(1, None),
                height=dp(48),
                radius=[12],
                elevation=0,
                md_bg_color=(0.90, 0.95, 1, 1) if is_sel else (0.97, 0.98, 0.99, 1),
                line_color=(0.1, 0.45, 0.85, 0.6) if is_sel else (0.89, 0.91, 0.94, 1),
                line_width=1.2 if is_sel else 1,
                padding=[dp(12), 0, dp(8), 0],
                spacing=dp(10),
                ripple_behavior=True,
                on_release=lambda x, c=code: [dlg_ref['dlg'].dismiss() if dlg_ref['dlg'] else None, on_select(c)],
            )
            card.add_widget(
                MDIcon(
                    icon=ico_name,
                    theme_text_color='Custom',
                    text_color=(0.1, 0.45, 0.85, 1) if is_sel else ico_col,
                    font_size='20sp',
                    size_hint=(None, None),
                    size=(dp(22), dp(22)),
                    pos_hint={'center_y': 0.5},
                )
            )
            card.add_widget(
                MDLabel(
                    text=label,
                    font_name='ArabicFont',
                    theme_text_color='Custom',
                    text_color=(0.08, 0.35, 0.72, 1) if is_sel else (0.15, 0.17, 0.22, 1),
                    font_style='Subtitle2',
                    bold=True,
                    halign='left',
                    valign='middle',
                    shorten=True,
                )
            )
            if is_sel:
                card.add_widget(
                    MDIcon(
                        icon='check-circle',
                        theme_text_color='Custom',
                        text_color=(0.1, 0.45, 0.85, 1),
                        font_size='20sp',
                        size_hint=(None, None),
                        size=(dp(22), dp(22)),
                        pos_hint={'center_y': 0.5},
                    )
                )
            if on_delete_item and opt.get('deletable', False):
                card.add_widget(
                    self._ui_icon_pill(
                        'trash-can-outline',
                        (1, 0.93, 0.93, 1),
                        (0.82, 0.14, 0.14, 1),
                        on_release=lambda x, c=code: on_delete_item(c),
                        size_dp=34,
                        icon_size='18sp',
                    )
                )
            lst.add_widget(card)

        scroll.add_widget(lst)
        content.add_widget(scroll)
        content.add_widget(
            self._ui_card_btn(
                'FERMER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                height=44,
                on_release=lambda *a: dlg_ref['dlg'].dismiss() if dlg_ref['dlg'] else None,
            )
        )

        dlg = MDDialog(
            title=title,
            type='custom',
            content_cls=content,
            size_hint=(0.90, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        dlg_ref['dlg'] = dlg
        self._active_filter_dialog = dlg
        self._fit_dialog(dlg)
        dlg.open()
        Clock.schedule_once(lambda dt: setattr(scroll, 'scroll_y', 1.0), 0.05)
        return dlg

    def _ui_confirm_dialog(
        self,
        title: str,
        message: str,
        confirm_text: str = 'CONFIRMER',
        cancel_text: str = 'ANNULER',
        icon: str = 'alert-circle-outline',
        confirm_icon: str = 'check-circle-outline',
        accent_color=(0.82, 0.12, 0.12, 1),
        soft_bg=(1, 0.94, 0.94, 1),
        border_color=(0.94, 0.78, 0.78, 1),
        on_confirm=None,
    ):
        content = MDBoxLayout(
            orientation='vertical',
            size_hint_y=None,
            adaptive_height=True,
            spacing=dp(12),
            padding=[0, 0, 0, dp(2)],
        )
        badge = self._ui_icon_pill(icon, soft_bg, accent_color, size_dp=56, radius=28, icon_size='28sp')
        badge.pos_hint = {'center_x': 0.5}
        content.add_widget(badge)

        content.add_widget(
            MDLabel(
                text=title,
                halign='center',
                font_style='H6',
                bold=True,
                theme_text_color='Custom',
                text_color=accent_color,
                size_hint_y=None,
                height=dp(28),
            )
        )

        msg_card = MDCard(
            orientation='vertical',
            size_hint=(1, None),
            adaptive_height=True,
            radius=[12],
            elevation=0,
            md_bg_color=soft_bg,
            line_color=border_color,
            line_width=1,
            padding=[dp(14), dp(10)],
        )
        msg_card.add_widget(
            self._ui_wrap_label(
                message,
                font_style='Body2',
                halign='center',
                bold=True,
                theme_color='Custom',
                text_color=(0.22, 0.16, 0.16, 1),
                min_h=22,
            )
        )
        content.add_widget(msg_card)

        dlg_ref = {'dlg': None}

        def _confirm_cb(*args):
            if dlg_ref['dlg']:
                dlg_ref['dlg'].dismiss()
            if on_confirm:
                on_confirm()

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(48))
        btns.add_widget(
            self._ui_card_btn(
                cancel_text,
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                size_hint_x=0.42,
                on_release=lambda x: dlg_ref['dlg'].dismiss() if dlg_ref['dlg'] else None,
            )
        )
        btns.add_widget(
            self._ui_card_btn(
                confirm_text,
                accent_color,
                text_color=(1, 1, 1, 1),
                icon=confirm_icon,
                size_hint_x=0.58,
                elevation=1,
                on_release=_confirm_cb,
            )
        )
        content.add_widget(btns)

        dlg = MDDialog(
            title='',
            type='custom',
            content_cls=content,
            size_hint=(0.90, None),
            radius=[20, 20, 20, 20],
            buttons=[],
        )
        dlg_ref['dlg'] = dlg
        dlg.open()
        return dlg

    def _preserve_rv_scroll_on_append(self, rv, new_data: list, reset: bool):
        if not rv:
            return

        rv._scroll_gen = getattr(rv, '_scroll_gen', 0) + 1
        gen = rv._scroll_gen

        if reset:
            rv._restoring_scroll = True
            rv.data = new_data
            rv.refresh_from_data()
            if getattr(rv, 'effect_y', None):
                rv.effect_y.velocity = 0.0
                rv.effect_y.value = 0.0
            rv.scroll_y = 1.0
            rv._restoring_scroll = False
            rv.loading_lock = False
            return

        if not new_data:
            rv.loading_lock = False
            return

        old_count = len(rv.data)
        old_sy = max(0.0, min(1.0, float(rv.scroll_y)))
        vp_h = max(1.0, float(rv.height))
        lm = getattr(rv, 'layout_manager', None)
        if lm:
            item_h = float(lm.default_size[1] or dp(88))
            sp = float(lm.spacing or 0)
            pad = lm.padding or [0, 0, 0, 0]
            pad_y = float(pad[1]) + float(pad[3]) if len(pad) >= 4 else 0.0
        else:
            item_h, sp, pad_y = float(dp(88)), float(dp(8)), 0.0

        def _calc_h(n: int) -> float:
            if n <= 0:
                return pad_y
            return pad_y + n * item_h + max(0, n - 1) * sp

        old_h = float(lm.height) if (lm and getattr(lm, 'height', 0) > 0) else _calc_h(old_count)
        dist_from_top = (1.0 - old_sy) * max(0.0, old_h - vp_h)

        rv._restoring_scroll = True
        rv.data.extend(new_data)
        est_new_h = _calc_h(len(rv.data))
        rv.refresh_from_data()

        def _apply_exact_scroll(*_):
            if getattr(rv, '_scroll_gen', 0) != gen:
                return
            cur_vp_h = max(1.0, float(rv.height))
            cur_lm = getattr(rv, 'layout_manager', None)
            actual_h = float(cur_lm.height) if (cur_lm and getattr(cur_lm, 'height', 0) > old_h) else est_new_h
            if actual_h <= cur_vp_h:
                target_sy = 1.0
                if getattr(rv, 'effect_y', None):
                    rv.effect_y.min = 0.0
                    rv.effect_y.max = 0.0
                    rv.effect_y.value = 0.0
                    rv.effect_y.velocity = 0.0
            else:
                scrollable_h = max(1.0, actual_h - cur_vp_h)
                clamped_dist = max(0.0, min(scrollable_h, dist_from_top))
                target_sy = max(0.0, min(1.0, 1.0 - (clamped_dist / scrollable_h)))
                if getattr(rv, 'effect_y', None):
                    rv.effect_y.min = -scrollable_h
                    rv.effect_y.max = 0.0
                    rv.effect_y.value = -clamped_dist
                    rv.effect_y.velocity = 0.0
            rv.scroll_y = target_sy

        def _release_scroll_lock(*_):
            if getattr(rv, '_scroll_gen', 0) != gen:
                return
            _apply_exact_scroll()
            rv._restoring_scroll = False
            rv.loading_lock = False

        _apply_exact_scroll()
        Clock.schedule_once(_apply_exact_scroll, 0)
        Clock.schedule_once(_release_scroll_lock, 0.06)

    # --------------------------------------------------------------------------
    # 7.1 CYCLE DE VIE, LICENCE & PERMISSIONS ANDROID
    # --------------------------------------------------------------------------
    def build(self):
        Builder.load_string(KV_BUILDER)
        self.title = AppConstants.APP_NAME
        self.theme_cls.primary_palette = 'Blue'
        self.theme_cls.accent_palette = 'Amber'
        self.theme_cls.theme_style = 'Light'
        for skey, spec in (
            ('H4', ['ArabicFont', 34, False, 0.25]),
            ('H5', ['ArabicFont', 24, False, 0]),
            ('H6', ['ArabicFont', 20, False, 0.15]),
            ('Subtitle1', ['ArabicFont', 16, False, 0.15]),
            ('Subtitle2', ['ArabicFont', 14, False, 0.1]),
            ('Body1', ['ArabicFont', 16, False, 0.5]),
            ('Body2', ['ArabicFont', 14, False, 0.25]),
            ('Button', ['ArabicFont', 14, True, 1.25]),
            ('Caption', ['ArabicFont', 12, False, 0.4]),
        ):
            self.theme_cls.font_styles[skey] = spec

        self.data_dir = self.user_data_dir
        self.db = DatabaseManager()
        self.is_seller_mode = self.db.get_bool_setting('config_seller_mode', default=False)

        self.root_box = MDBoxLayout(orientation='vertical')
        self.sm = MDScreenManager()
        self.sm.add_widget(self._build_login_screen())
        self.sm.add_widget(self._build_dashboard_screen())
        self.sm.add_widget(self._build_products_screen())
        self.sm.add_widget(self._build_cart_screen())
        self.root_box.add_widget(self.sm)

        self.status_bar_bg = MDCard(
            orientation='horizontal',
            size_hint_y=None,
            height=dp(40),
            radius=[0],
            md_bg_color=(0.14, 0.16, 0.20, 1),
            line_color=(0.24, 0.27, 0.33, 1),
            line_width=1,
            elevation=0,
            padding=[dp(14), 0],
            spacing=dp(8),
        )
        self.status_bar_bg.add_widget(MDBoxLayout())
        self.status_bar_icon = MDIcon(
            icon='shield-check-outline',
            theme_text_color='Custom',
            text_color=(0.75, 0.85, 0.98, 1),
            font_size='17sp',
            size_hint=(None, None),
            size=(dp(18), dp(18)),
            pos_hint={'center_y': 0.5},
        )
        self.status_bar_label = MDLabel(
            text='Prêt',
            halign='center',
            valign='middle',
            theme_text_color='Custom',
            text_color=(1, 1, 1, 1),
            font_style='Caption',
            bold=True,
            adaptive_width=True,
            pos_hint={'center_y': 0.5},
        )
        self.status_bar_bg.add_widget(self.status_bar_icon)
        self.status_bar_bg.add_widget(self.status_bar_label)
        self.status_bar_bg.add_widget(MDBoxLayout())
        self.root_box.add_widget(self.status_bar_bg)

        Window.bind(on_keyboard=self.on_keyboard)
        return self.root_box

    def on_start(self):
        self.request_android_permissions()
        if platform == 'android' and ToneGenerator is not None:
            try:
                self.tone_gen = ToneGenerator(3, 100)
            except Exception:
                self.tone_gen = None
        Clock.schedule_once(self._deferred_start, 0.4)

    def on_stop(self):
        try:
            self.trigger_auto_backups()
            if self.db:
                self.db.cleanup_wal_files()
        except Exception:
            pass

    def request_android_permissions(self):
        if platform != 'android':
            return
        try:
            from android.permissions import Permission, request_permissions
            from jnius import autoclass

            version_cls = autoclass('android.os.Build$VERSION')
            perms = [
                Permission.BLUETOOTH,
                Permission.BLUETOOTH_ADMIN,
                Permission.ACCESS_COARSE_LOCATION,
                Permission.ACCESS_FINE_LOCATION,
            ]
            if version_cls.SDK_INT <= 32:
                perms.extend([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE])
            if version_cls.SDK_INT >= 31:
                perms.extend(['android.permission.BLUETOOTH_CONNECT', 'android.permission.BLUETOOTH_SCAN'])
            request_permissions(perms, lambda p, r: None)
        except Exception:
            pass

    def play_sound(self, sound_type: str):
        if platform == 'android' and self.tone_gen:
            try:
                if sound_type == 'success':
                    self.tone_gen.startTone(24, 150)
                elif sound_type == 'error':
                    self.tone_gen.startTone(97, 300)
                elif sound_type == 'duplicate':
                    self.tone_gen.startTone(29, 150)
            except Exception:
                pass

    def notify(self, text: str, msg_type: str = 'info'):
        if threading.current_thread() is threading.main_thread():
            self._apply_notify_ui(str(text), msg_type)
        else:
            Clock.schedule_once(lambda dt, t=str(text), m=msg_type: self._apply_notify_ui(t, m), 0)

    def _apply_notify_ui(self, text: str, msg_type: str):
        if not self.status_bar_label or not self.status_bar_bg:
            return
        styles = {
            'success': ((0.05, 0.56, 0.26, 1), (0.04, 0.44, 0.20, 1), 'check-circle-outline'),
            'error': ((0.80, 0.12, 0.14, 1), (0.64, 0.08, 0.10, 1), 'alert-circle-outline'),
            'warning': ((0.88, 0.46, 0.02, 1), (0.72, 0.36, 0.0, 1), 'alert-outline'),
            'info': ((0.14, 0.36, 0.65, 1), (0.10, 0.28, 0.54, 1), 'information-outline'),
        }
        bg_c, line_c, ico_name = styles.get(msg_type, styles['info'])
        Animation.cancel_all(self.status_bar_bg)
        self.status_bar_label.text = text
        if self.status_bar_icon:
            self.status_bar_icon.icon = ico_name
            self.status_bar_icon.text_color = (1, 1, 1, 1)
        Animation(md_bg_color=bg_c, line_color=line_c, duration=0.18).start(self.status_bar_bg)
        if self._notify_event:
            self._notify_event.cancel()
        self._notify_event = Clock.schedule_once(self._reset_notification_state, 3.0)

    def _reset_notification_state(self, dt):
        if self.status_bar_label and self.status_bar_bg:
            self._notify_event = None
            self.status_bar_label.text = 'Prêt'
            if self.status_bar_icon:
                self.status_bar_icon.icon = 'shield-check-outline'
                self.status_bar_icon.text_color = (0.75, 0.85, 0.98, 1)
            Animation.cancel_all(self.status_bar_bg)
            Animation(
                md_bg_color=(0.14, 0.16, 0.20, 1),
                line_color=(0.24, 0.27, 0.33, 1),
                duration=0.28,
            ).start(self.status_bar_bg)

    def on_keyboard(self, window, key, scancode, codepoint, modifier):
        if key == 27:
            if getattr(self, 'scan_dialog', None):
                self.close_barcode_scanner()
                return True
            act_dlg = getattr(self, 'activation_dialog_ref', None)
            if act_dlg and act_dlg.parent:
                if getattr(self, '_activation_trial_expired', True):
                    return True
                self.continue_trial(act_dlg)
                return True
            if getattr(self, '_import_in_progress', False) or getattr(self, '_restore_in_progress', False) or getattr(self, '_backup_in_progress', False):
                return True
            win_children = list(Window.children)
            if win_children and isinstance(win_children[0], MDDatePicker):
                win_children[0].dismiss()
                return True
            dialog_names = (
                'exp_cat_dialog', '_active_filter_dialog', 'hist_filter_dialog',
                'del_exp_dialog', 'confirm_del_dialog', 'smart_pdf_dialog', 'pdf_ready_dialog',
                'final_restore_confirm', 'stock_error_dialog', 'warn_cost_dlg',
                'confirm_del_fam_dialog', 'add_fam_dialog', 'del_conf_dialog',
                'debt_dialog', 'overpay_dialog',
                'cat_dialog', 'ae_dialog', 'options_dialog', 'srv_dialog',
                'edit_dialog', 'pay_dialog', 'simple_pay_dialog',
                'zoom_dialog', 'img_picker_dialog', 'bt_dialog', 'family_dialog',
                'filter_dialog', 'entity_dialog', 'entity_hist_dialog', 'pending_dialog',
                'history_exp_dialog', 'expense_dialog', 'mgmt_dialog',
                'store_settings_dialog', 'auth_dialog', 'toggle_dialog',
                'settings_menu_dialog', 'pass_dialog', 'logout_diag',
                'mapping_diag', 'import_diag', 'import_report_dialog',
                'restore_dialog', 'dialog',
            )
            open_dialogs = []
            seen_ids = set()
            for idx, d_name in enumerate(dialog_names):
                dlg = getattr(self, d_name, None)
                if dlg and dlg.parent and id(dlg) not in seen_ids:
                    seen_ids.add(id(dlg))
                    z_idx = win_children.index(dlg) if dlg in win_children else (10000 + idx)
                    open_dialogs.append((z_idx, d_name, dlg))
            if open_dialogs:
                open_dialogs.sort(key=lambda item: item[0])
                _, top_name, top_dlg = open_dialogs[0]
                if top_name == 'simple_pay_dialog':
                    self.editing_tx_id = None
                    self.current_editing_date = None
                    self.temp_note = ''
                top_dlg.dismiss()
                return True
            cur = self.sm.current
            if cur == 'cart':
                self.back_to_products()
                return True
            if cur == 'products':
                self.go_back()
                return True
            if cur == 'dashboard':
                self.logout()
                return True
        return False

    # --- Gestion Licence / Essai ---
    def get_device_id(self) -> str:
        if platform == 'android':
            try:
                resolver = PythonActivity.mActivity.getContentResolver()
                from jnius import autoclass
                secure = autoclass('android.provider.Settings$Secure')
                aid = secure.getString(resolver, secure.ANDROID_ID)
                return str(aid) if aid else 'ANDROID_UNKNOWN'
            except Exception:
                return 'ANDROID_ERR_ID'
        if platform == 'win':
            return 'PC_DEBUG_ID_12345'
        return 'UNKNOWN_DEVICE_ID'

    def _get_hidden_sys_file(self) -> str:
        if platform == 'android':
            try:
                from jnius import autoclass
                env = autoclass('android.os.Environment')
                pub_dir = env.getExternalStoragePublicDirectory(env.DIRECTORY_DOCUMENTS).getAbsolutePath()
                return os.path.join(pub_dir, '.magpro_mobile_sys.dat')
            except Exception:
                return os.path.join(self.user_data_dir, '.magpro_mobile_sys.dat')
        return os.path.join(os.path.expanduser('~'), '.magpro_mobile_sys.dat')

    def load_persistent_data(self) -> dict:
        candidates = []
        ext_path = self._get_hidden_sys_file()
        int_path = os.path.join(self.user_data_dir, '.magpro_sys_cache.dat')

        for p in (ext_path, int_path):
            if p and os.path.exists(p):
                try:
                    with open(p, 'r', encoding='utf-8') as f:
                        d = json.load(f)
                        if isinstance(d, dict):
                            candidates.append(d)
                except Exception:
                    pass

        if getattr(self, 'db', None):
            try:
                raw_db = self.db.get_setting('sys_trial_state', '')
                if raw_db:
                    d_db = json.loads(raw_db)
                    if isinstance(d_db, dict):
                        candidates.append(d_db)
            except Exception:
                pass

        if not candidates:
            return {'trial_start': None, 'attempts': 0}

        earliest_dt = None
        earliest_str = None
        max_attempts = 0
        for c in candidates:
            try:
                max_attempts = max(max_attempts, int(c.get('attempts', 0) or 0))
            except Exception:
                pass
            ts_str = c.get('trial_start')
            if ts_str:
                try:
                    dt_obj = datetime.strptime(str(ts_str)[:19], '%Y-%m-%d %H:%M:%S')
                    if earliest_dt is None or dt_obj < earliest_dt:
                        earliest_dt = dt_obj
                        earliest_str = dt_obj.strftime('%Y-%m-%d %H:%M:%S')
                except Exception:
                    pass

        merged = {'trial_start': earliest_str, 'attempts': max_attempts}
        if len(candidates) < 3 and earliest_str is not None:
            self.save_persistent_data(merged)
        return merged

    def save_persistent_data(self, data: dict):
        ext_path = self._get_hidden_sys_file()
        int_path = os.path.join(self.user_data_dir, '.magpro_sys_cache.dat')

        for p in (ext_path, int_path):
            if not p:
                continue
            try:
                parent = os.path.dirname(p)
                if parent:
                    os.makedirs(parent, exist_ok=True)
                with open(p, 'w', encoding='utf-8') as f:
                    json.dump(data, f)
            except Exception:
                pass

        if getattr(self, 'db', None):
            try:
                self.db.set_setting('sys_trial_state', json.dumps(data))
            except Exception:
                pass

    def check_license_validity(self) -> tuple:
        try:
            stored_key = self.db.get_setting('license_key')
            device_id = self.get_device_id()
            expected_key = hashlib.sha256(f'magpro_mobile_v6_{device_id}_secure_key'.encode()).hexdigest()
            if stored_key and stored_key == expected_key:
                return 'ACTIVATED', 0

            now = datetime.now()
            p_data = self.load_persistent_data()
            install_str = p_data.get('trial_start')
            if not install_str:
                install_str = now.strftime('%Y-%m-%d %H:%M:%S')
                p_data['trial_start'] = install_str
                self.save_persistent_data(p_data)
            try:
                install_date = datetime.strptime(install_str, '%Y-%m-%d %H:%M:%S')
            except Exception:
                install_date = now
            days_used = (now - install_date).days
            if days_used < 0:
                return 'TAMPERED', 0
            if days_used <= 7:
                return 'TRIAL', 7 - days_used
            return 'EXPIRED', 0
        except Exception:
            return 'EXPIRED', 0

    def _deferred_start(self, dt):
        status, days_left = self.check_license_validity()
        if status in ('EXPIRED', 'TAMPERED'):
            self.show_activation_dialog(trial_expired=True)
        elif status == 'TRIAL':
            self.show_activation_dialog(trial_expired=False, days_left=days_left)
        else:
            self.continue_trial(None)

    def continue_trial(self, dialog_ref):
        if dialog_ref:
            dialog_ref.dismiss()
        self.activation_dialog_ref = None
        self._activation_trial_expired = False
        if hasattr(self, 'password_field'):
            self.password_field.set_raw_value('')
        self.update_dashboard_layout()
        Clock.unschedule(self.trigger_auto_backups)
        Clock.schedule_interval(self.trigger_auto_backups, 3600)
        self.trigger_auto_backups()

    def validate_activation(self, key_input: str, dialog_ref):
        device_id = self.get_device_id()
        expected = hashlib.sha256(f'magpro_mobile_v6_{device_id}_secure_key'.encode()).hexdigest()
        if str(key_input or '').strip().lower() == expected:
            self.db.set_setting('license_key', expected)
            self.notify('Activation réussie ! Bienvenue.', 'success')
            if dialog_ref:
                dialog_ref.dismiss()
            self.activation_dialog_ref = None
            self._activation_trial_expired = False
            self.continue_trial(None)
        else:
            self.notify('Clé invalide. Veuillez vérifier.', 'error')

    def show_activation_dialog(self, trial_expired: bool = True, days_left: int = 0):
        self._activation_trial_expired = bool(trial_expired)
        device_id = self.get_device_id()

        wrapper = MDBoxLayout(
            orientation='vertical',
            size_hint_y=None,
            height=min(dp(560), Window.height * 0.84),
            padding=0,
        )
        scroll = MDScrollView(size_hint_y=1, bar_width=dp(3))
        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(11),
            size_hint_y=None,
            adaptive_height=True,
            padding=[0, dp(2), 0, dp(4)],
        )

        icon_name = 'shield-lock' if trial_expired else 'shield-check'
        icon_col = (0.82, 0.12, 0.12, 1) if trial_expired else (0.09, 0.44, 0.84, 1)
        badge_bg = (1, 0.92, 0.92, 1) if trial_expired else (0.90, 0.95, 1, 1)

        badge = self._ui_icon_pill(icon_name, badge_bg, icon_col, size_dp=66, radius=33, icon_size='34sp')
        badge.pos_hint = {'center_x': 0.5}
        content.add_widget(badge)

        if trial_expired:
            t_txt = "Période d'essai terminée"
            s_txt = "L'application est verrouillée.\nVeuillez saisir votre clé d'activation pour continuer."
            t_col = (0.80, 0.12, 0.12, 1)
        else:
            t_txt = f"Essai : {days_left} jour(s) restant(s)"
            s_txt = "Profitez de votre essai gratuit ou activez l'application dès maintenant."
            t_col = (0.09, 0.44, 0.84, 1)

        hdr_box = MDBoxLayout(orientation='vertical', spacing=dp(4), adaptive_height=True)
        hdr_box.add_widget(MDLabel(text=t_txt, halign='center', font_style='H6', bold=True, theme_text_color='Custom', text_color=t_col, adaptive_height=True))
        hdr_box.add_widget(MDLabel(text=s_txt, halign='center', font_style='Caption', theme_text_color='Secondary', adaptive_height=True))
        content.add_widget(hdr_box)

        id_card = MDCard(
            orientation='vertical',
            radius=[12],
            padding=[dp(12), dp(10)],
            md_bg_color=(0.95, 0.97, 1, 1),
            line_color=(0.82, 0.88, 0.96, 1),
            line_width=1,
            elevation=0,
            size_hint_y=None,
            adaptive_height=True,
            spacing=dp(4),
        )
        id_card.add_widget(MDLabel(text="ID de l'appareil :", halign='left', font_style='Caption', bold=True, theme_text_color='Secondary', adaptive_height=True))
        id_row = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint_y=None, height=dp(40))
        id_row.add_widget(MDTextField(text=device_id, readonly=True, font_size='15sp', mode='line', active_line=False, line_color_normal=(0, 0, 0, 0), size_hint_x=1, pos_hint={'center_y': 0.5}))
        id_row.add_widget(
            self._ui_icon_pill(
                'content-copy',
                (0.86, 0.92, 1, 1),
                (0.09, 0.44, 0.84, 1),
                on_release=lambda x: [Clipboard.copy(device_id), self.notify('ID copié', 'success')],
                size_dp=38,
                icon_size='19sp',
            )
        )
        id_card.add_widget(id_row)
        content.add_widget(id_card)

        key_card = MDCard(
            orientation='horizontal',
            size_hint=(1, None),
            height=dp(58),
            radius=[12],
            elevation=0,
            md_bg_color=(0.97, 0.98, 0.99, 1),
            line_color=(0.86, 0.88, 0.93, 1),
            line_width=1,
            padding=[dp(12), 0, dp(8), 0],
            spacing=dp(8),
        )
        self.field_key = NoMenuTextField(
            hint_text='Saisir la clé de licence',
            mode='line',
            line_color_normal=(0, 0, 0, 0),
            line_color_focus=(0, 0, 0, 0),
            size_hint_x=1,
            pos_hint={'center_y': 0.5},
        )
        key_card.add_widget(self.field_key)
        key_card.add_widget(
            self._ui_icon_pill(
                'content-paste',
                (0.91, 0.95, 1, 1),
                (0.09, 0.44, 0.84, 1),
                on_release=lambda x: setattr(self.field_key, 'text', Clipboard.paste()),
                size_dp=38,
                icon_size='19sp',
            )
        )
        content.add_widget(key_card)

        content.add_widget(
            self._ui_card_btn(
                "ACTIVER L'APPLICATION",
                (0.0, 0.62, 0.25, 1),
                icon='check-decagram',
                height=50,
                font_size='15sp',
                elevation=1,
                on_release=lambda x: self.validate_activation(self.field_key.text, self.activation_dialog_ref),
            )
        )

        class _OnlineActCard(MDCard):
            text = StringProperty("Demander l'activation en ligne")
            icon = StringProperty('earth')

            def __init__(self, **kwargs):
                super().__init__(**kwargs)
                self.orientation = 'horizontal'
                self.size_hint = (1, None)
                self.height = dp(48)
                self.radius = [12]
                self.elevation = 0
                self.md_bg_color = (0.09, 0.44, 0.84, 1)
                self.ripple_behavior = True
                self.spacing = dp(8)
                self.padding = [dp(10), 0]

                self._ico = MDIcon(icon=self.icon, theme_text_color='Custom', text_color=(1, 1, 1, 1), font_size='20sp', size_hint_x=None, width=dp(22), pos_hint={'center_y': 0.5})
                self._lbl = MDLabel(text=self.text, bold=True, font_size='14sp', theme_text_color='Custom', text_color=(1, 1, 1, 1), adaptive_width=True, shorten=True, pos_hint={'center_y': 0.5})
                self.add_widget(MDBoxLayout())
                self.add_widget(self._ico)
                self.add_widget(self._lbl)
                self.add_widget(MDBoxLayout())
                self.bind(
                    text=lambda inst, val: setattr(self._lbl, 'text', val),
                    icon=lambda inst, val: setattr(self._ico, 'icon', val),
                    disabled=lambda inst, val: setattr(self, 'opacity', 0.65 if val else 1.0),
                )

        self.online_act_btn = _OnlineActCard(on_release=self.request_online_activation_mobile)
        content.add_widget(self.online_act_btn)

        if not trial_expired:
            content.add_widget(
                self._ui_card_btn(
                    "Continuer l'essai gratuit",
                    (0.92, 0.94, 0.97, 1),
                    text_color=(0.09, 0.44, 0.84, 1),
                    height=44,
                    font_size='13.5sp',
                    on_release=lambda x: self.continue_trial(self.activation_dialog_ref),
                )
            )

        scroll.add_widget(content)
        wrapper.add_widget(scroll)

        self.activation_dialog_ref = MDDialog(
            title='',
            type='custom',
            content_cls=wrapper,
            size_hint=(0.94, None),
            auto_dismiss=False,
            radius=[20, 20, 20, 20],
            buttons=[],
        )
        self._fit_dialog(self.activation_dialog_ref, top_pad=14, side_pad=14, bottom_pad=14, spacer_top=0)
        self.activation_dialog_ref.open()

    def request_online_activation_mobile(self, instance=None):
        if self.online_act_btn.disabled:
            return
        p_data = self.load_persistent_data()
        attempts = p_data.get('attempts', 0)
        if attempts >= 5:
            self.notify('Limite atteinte (5). Fonctionnalité bloquée.', 'error')
            self.online_act_btn.text = 'Bloqué définitivement'
            self.online_act_btn.icon = 'cancel'
            self.online_act_btn.disabled = True
            return
        self.online_act_btn.text = 'Vérification...'
        self.online_act_btn.icon = 'sync'
        self.online_act_btn.disabled = True
        threading.Thread(target=self._online_activation_worker_mobile, daemon=True).start()

    def _online_activation_worker_mobile(self):
        BOT_TOKEN = "GITHUB_SECRET_TELEGRAM_BOT_TOKEN"
        MASTER_ADMIN_ID = "GITHUB_SECRET_TELEGRAM_ADMIN_ID"
        device_id = self.get_device_id()

        if "GITHUB_SECRET" in BOT_TOKEN or "GITHUB_SECRET" in MASTER_ADMIN_ID:
            def not_configured(dt):
                self.notify('❌ Service non configuré dans cette version.', 'error')
                self.online_act_btn.text = "Demander l'activation en ligne"
                self.online_act_btn.icon = 'earth'
                self.online_act_btn.disabled = False
            Clock.schedule_once(not_configured, 0)
            return

        try:
            import requests
            requests.get('https://api.telegram.org', timeout=5)
        except Exception:
            def no_internet(dt):
                self.notify('❌ Pas de connexion Internet', 'error')
                self.online_act_btn.text = "Demander l'activation en ligne"
                self.online_act_btn.icon = 'earth'
                self.online_act_btn.disabled = False
            Clock.schedule_once(no_internet, 0)
            return

        Clock.schedule_once(lambda dt: setattr(self.online_act_btn, 'text', 'Envoi de la demande...'), 0)
        msg = f'📱 **طلب تفعيل تطبيق MagPro Mobile**\n\n💻 **رقم الجهاز (ID):** `{device_id}`'
        reply_markup = {
            "inline_keyboard": [
                [{"text": "✅ تفعيل التطبيق", "callback_data": f"ACT_OK|MOB|{device_id}"}],
                [{"text": "❌ رفض الطلب", "callback_data": f"ACT_NO|MOB|{device_id}"}]
            ]
        }
        try:
            send_resp = requests.post(
                f'https://api.telegram.org/bot{BOT_TOKEN}/sendMessage',
                json={'chat_id': MASTER_ADMIN_ID, 'text': msg, 'parse_mode': 'Markdown', 'reply_markup': reply_markup},
                timeout=10,
            )
            if send_resp.status_code != 200:
                raise RuntimeError(f'HTTP {send_resp.status_code}')
        except Exception:
            def no_server(dt):
                self.notify("❌ Échec d'envoi au serveur", 'error')
                self.online_act_btn.text = "Demander l'activation en ligne"
                self.online_act_btn.icon = 'earth'
                self.online_act_btn.disabled = False
            Clock.schedule_once(no_server, 0)
            return

        def _count_attempt(dt):
            p_cur = self.load_persistent_data()
            p_cur['attempts'] = int(p_cur.get('attempts', 0) or 0) + 1
            self.save_persistent_data(p_cur)
        Clock.schedule_once(_count_attempt, 0)

        Clock.schedule_once(lambda dt: setattr(self.online_act_btn, 'text', 'En attente du dev... (Ne pas fermer)'), 0)
        offset = None
        try:
            resp = requests.get(f'https://api.telegram.org/bot{BOT_TOKEN}/getUpdates', timeout=10).json()
            if resp.get('result'):
                offset = resp['result'][-1]['update_id'] + 1
        except Exception:
            pass

        timeout_end = time.time() + 900
        is_activated = False
        while time.time() < timeout_end:
            if not getattr(self, 'activation_dialog_ref', None):
                return
            try:
                params = {'timeout': 10, 'allowed_updates': ['message', 'callback_query']}
                if offset:
                    params['offset'] = offset
                resp = requests.get(f'https://api.telegram.org/bot{BOT_TOKEN}/getUpdates', params=params, timeout=15)
                if resp.status_code == 200:
                    data = resp.json()
                    for result in data.get('result', []):
                        offset = result['update_id'] + 1
                        if 'callback_query' in result:
                            cb = result['callback_query']
                            cb_data = cb.get('data', '')
                            cb_id = cb.get('id')
                            msg_id = cb.get('message', {}).get('message_id')
                            parts = cb_data.split('|')
                            if len(parts) == 3 and parts[1] == 'MOB' and parts[2] == device_id:
                                if parts[0] == 'ACT_OK':
                                    salt = f'magpro_mobile_v6_{device_id}_secure_key'
                                    expected_key = hashlib.sha256(salt.encode()).hexdigest().lower()
                                    requests.post(
                                        f'https://api.telegram.org/bot{BOT_TOKEN}/answerCallbackQuery',
                                        json={'callback_query_id': cb_id, 'text': 'تم تفعيل التطبيق ✅'},
                                    )
                                    requests.post(
                                        f'https://api.telegram.org/bot{BOT_TOKEN}/editMessageText',
                                        json={'chat_id': MASTER_ADMIN_ID, 'message_id': msg_id, 'text': msg + '\n\n✅ **الحالة: تم التفعيل بنجاح**', 'parse_mode': 'Markdown'},
                                    )
                                    is_activated = True

                                    def apply_activation_ui(dt, ek=expected_key):
                                        if hasattr(self, 'field_key'):
                                            self.field_key.text = ek
                                        if hasattr(self, 'online_act_btn'):
                                            self.online_act_btn.text = 'Activé avec succès!'
                                            self.online_act_btn.icon = 'check-circle'
                                        self.notify('Activation réussie !', 'success')
                                        self.validate_activation(ek, self.activation_dialog_ref)

                                    Clock.schedule_once(apply_activation_ui, 0)
                                    return
                                elif parts[0] == 'ACT_NO':
                                    requests.post(
                                        f'https://api.telegram.org/bot{BOT_TOKEN}/answerCallbackQuery',
                                        json={'callback_query_id': cb_id, 'text': 'تم الرفض ❌'},
                                    )
                                    requests.post(
                                        f'https://api.telegram.org/bot{BOT_TOKEN}/editMessageText',
                                        json={'chat_id': MASTER_ADMIN_ID, 'message_id': msg_id, 'text': msg + '\n\n❌ **الحالة: تم الرفض**', 'parse_mode': 'Markdown'},
                                    )

                                    def apply_reject_ui(dt):
                                        self.notify("❌ Demande refusée par l'administrateur.", 'error')
                                        if hasattr(self, 'online_act_btn'):
                                            self.online_act_btn.text = "Demander l'activation en ligne"
                                            self.online_act_btn.icon = 'earth'
                                            self.online_act_btn.disabled = False

                                    Clock.schedule_once(apply_reject_ui, 0)
                                    return

                        message = result.get('message', {})
                        text = str(message.get('text', '')).strip()
                        chat_id = str(message.get('chat', {}).get('id', ''))
                        if chat_id == MASTER_ADMIN_ID and text and len(text) >= 32:
                            received_key = text.strip().lower()
                            salt = f'magpro_mobile_v6_{device_id}_secure_key'
                            expected_key = hashlib.sha256(salt.encode()).hexdigest().lower()
                            if received_key == expected_key:
                                is_activated = True
                                requests.post(
                                    f'https://api.telegram.org/bot{BOT_TOKEN}/sendMessage',
                                    json={'chat_id': MASTER_ADMIN_ID, 'text': f'✅ **تم تفعيل تطبيق الهاتف بنجاح!**\n💻 رقم الجهاز: `{device_id}`', 'parse_mode': 'Markdown'},
                                )

                                def apply_activation_ui_manual(dt, rk=text.strip()):
                                    if hasattr(self, 'field_key'):
                                        self.field_key.text = rk
                                    if hasattr(self, 'online_act_btn'):
                                        self.online_act_btn.text = 'Activé avec succès!'
                                        self.online_act_btn.icon = 'check-circle'
                                    self.notify('Activation réussie !', 'success')
                                    self.validate_activation(rk, self.activation_dialog_ref)

                                Clock.schedule_once(apply_activation_ui_manual, 0)
                                return
            except Exception:
                pass
            time.sleep(3)

        if not is_activated:
            def timeout_ui(dt):
                self.notify("⚠️ Délai d'attente expiré.", 'warning')
                if hasattr(self, 'online_act_btn'):
                    self.online_act_btn.text = "Demander l'activation en ligne"
                    self.online_act_btn.icon = 'earth'
                    self.online_act_btn.disabled = False
            Clock.schedule_once(timeout_ui, 0)

    # --------------------------------------------------------------------------
    # 7.2 CONNEXION & MOT DE PASSE
    # --------------------------------------------------------------------------
    def _build_login_screen(self):
        screen = MDScreen(name='login')
        screen.md_bg_color = (0.94, 0.95, 0.97, 1)
        layout = MDFloatLayout()

        card = MDCard(
            orientation='vertical',
            size_hint=(0.88, None),
            adaptive_height=True,
            pos_hint={'center_x': 0.5, 'center_y': 0.53},
            padding=[dp(24), dp(28), dp(24), dp(24)],
            spacing=dp(14),
            radius=[24, 24, 24, 24],
            elevation=3,
            md_bg_color=(1, 1, 1, 1),
        )

        icon_badge = self._ui_icon_pill(
            'store-check',
            (0.91, 0.95, 1, 1),
            (0.1, 0.45, 0.85, 1),
            size_dp=78,
            radius=39,
            icon_size='44sp',
        )
        icon_badge.pos_hint = {'center_x': 0.5}
        card.add_widget(icon_badge)

        title_box = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True, spacing=dp(4))
        title_box.add_widget(MDLabel(text=AppConstants.APP_NAME, halign='center', font_style='H5', bold=True, adaptive_height=True, theme_text_color='Custom', text_color=(0.12, 0.14, 0.18, 1)))
        title_box.add_widget(MDLabel(text='Authentification', halign='center', font_style='Caption', theme_text_color='Secondary', adaptive_height=True))
        card.add_widget(title_box)

        pass_card = MDCard(
            orientation='horizontal',
            size_hint=(1, None),
            height=dp(58),
            radius=[12],
            elevation=0,
            md_bg_color=(0.95, 0.96, 0.98, 1),
            line_color=(0.85, 0.88, 0.94, 1),
            line_width=1,
            padding=[dp(14), dp(2), dp(4), dp(2)],
            spacing=dp(8),
        )
        pass_card.add_widget(
            MDIcon(
                icon='lock-outline',
                theme_text_color='Custom',
                text_color=(0.1, 0.45, 0.85, 1),
                font_size='22sp',
                size_hint_x=None,
                width=dp(24),
                pos_hint={'center_y': 0.5},
            )
        )
        self.password_field = SmartTextField(
            hint_text='Mot de passe',
            password=True,
            mode='line',
            active_line=False,
            line_color_normal=(0, 0, 0, 0),
            line_color_focus=(0, 0, 0, 0),
            pos_hint={'center_y': 0.5},
        )
        self.password_field.bind(on_text_validate=self.do_login)
        pass_card.add_widget(self.password_field)
        pass_card.add_widget(
            MDIconButton(
                icon='eye-off-outline',
                theme_text_color='Secondary',
                icon_size='20sp',
                pos_hint={'center_y': 0.5},
                on_release=lambda b: [setattr(self.password_field, 'password', not self.password_field.password), setattr(b, 'icon', 'eye-off-outline' if self.password_field.password else 'eye-outline')],
            )
        )
        card.add_widget(pass_card)

        card.add_widget(
            self._ui_card_btn(
                'SE CONNECTER',
                (0.1, 0.45, 0.85, 1),
                icon='login',
                height=50,
                font_size='16sp',
                elevation=2,
                on_release=self.do_login,
            )
        )
        layout.add_widget(card)

        btn_chg_pw = self._ui_card_btn(
            'Changer / Définir le mot de passe',
            (0.89, 0.91, 0.95, 1),
            text_color=(0.25, 0.3, 0.4, 1),
            icon='key-change',
            height=42,
            radius=21,
            font_size='12.5sp',
            border_color=(0.82, 0.85, 0.91, 1),
            on_release=self.show_change_password_dialog_login,
        )
        btn_chg_pw.size_hint = (None, None)
        btn_chg_pw.height = dp(42)
        btn_chg_pw.pos_hint = {'center_x': 0.5, 'y': 0.09}

        def _fit_chg_pw_btn(*_):
            btn_chg_pw._lbl.text_size = (None, None)
            btn_chg_pw._lbl.texture_update()
            needed = btn_chg_pw._lbl.texture_size[0] + dp(56)
            btn_chg_pw.width = max(dp(260), min(Window.width * 0.88, needed))

        Window.bind(width=lambda *_: _fit_chg_pw_btn())
        Clock.schedule_once(lambda dt: _fit_chg_pw_btn(), 0)
        layout.add_widget(btn_chg_pw)

        layout.add_widget(
            MDLabel(
                text=f'{AppConstants.APP_NAME} v{AppConstants.APP_VERSION} © {AppConstants.COPYRIGHT_YEAR}',
                halign='center',
                pos_hint={'center_x': 0.5, 'y': 0.02},
                size_hint_y=None,
                height=dp(20),
                font_style='Caption',
                theme_text_color='Hint',
            )
        )
        screen.add_widget(layout)
        return screen

    def do_login(self, x):
        self.notify('Connexion en cours...', 'info')
        p = self.password_field.get_value()
        user = self.db.login(p)
        if user:
            self.current_user_name = user.get('username', AppConstants.ADMIN_USERNAME)
            self.sm.current = 'dashboard'
            self.check_and_load_stats()
            self.notify('Bienvenue dans MagPro Mobile', 'success')
        else:
            self.play_sound('error')
            self.notify('Mot de passe incorrect', 'error')

    def show_change_password_dialog_login(self, instance):
        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(10),
            size_hint_y=None,
            adaptive_height=True,
            padding=[0, dp(4), 0, 0],
        )

        def _make_pw_box(hint_txt: str):
            card = MDCard(
                orientation='horizontal',
                size_hint=(1, None),
                height=dp(60),
                radius=[12],
                elevation=0,
                md_bg_color=(0.96, 0.97, 0.99, 1),
                line_color=(0.86, 0.88, 0.93, 1),
                line_width=1,
                padding=[dp(12), dp(2), dp(4), dp(2)],
                spacing=dp(8),
            )
            card.add_widget(MDIcon(icon='lock-outline', theme_text_color='Custom', text_color=(0.09, 0.44, 0.84, 1), font_size='20sp', size_hint_x=None, width=dp(22), pos_hint={'center_y': 0.5}))
            tf = SmartTextField(
                hint_text=hint_txt,
                password=True,
                mode='line',
                active_line=False,
                line_color_normal=(0, 0, 0, 0),
                line_color_focus=(0, 0, 0, 0),
                pos_hint={'center_y': 0.5},
            )
            card.add_widget(tf)
            card.add_widget(
                MDIconButton(
                    icon='eye-off-outline',
                    theme_text_color='Secondary',
                    icon_size='20sp',
                    pos_hint={'center_y': 0.5},
                    on_release=lambda b, f=tf: [setattr(f, 'password', not f.password), setattr(b, 'icon', 'eye-off-outline' if f.password else 'eye-outline')],
                )
            )
            return card, tf

        c_old, old_pass = _make_pw_box('Mot de passe actuel')
        c_new, new_pass = _make_pw_box('Nouveau mot de passe')
        c_cnf, conf_pass = _make_pw_box('Confirmer le mot de passe')
        content.add_widget(c_old)
        content.add_widget(c_new)
        content.add_widget(c_cnf)

        def _confirm(x):
            old_pass.error = False
            new_pass.error = False
            conf_pass.error = False
            cur_p = old_pass.get_value()
            new_p = new_pass.get_value()
            cnf_p = conf_pass.get_value()
            if not new_p:
                new_pass.error = True
                self.notify('Le nouveau mot de passe est vide.', 'error')
                return
            if new_p != cnf_p:
                new_pass.error = True
                conf_pass.error = True
                self.notify('Les mots de passe ne correspondent pas.', 'error')
                return
            if self.db.login(cur_p):
                self.db.update_admin_password(new_p)
                self.password_field.set_raw_value('')
                self.notify('Mot de passe modifié avec succès.', 'success')
                self.pass_dialog.dismiss()
            else:
                old_pass.error = True
                self.notify('Mot de passe actuel incorrect.', 'error')

        conf_pass.bind(on_text_validate=_confirm)

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(48), padding=[0, dp(4), 0, 0])
        btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.40, font_size='13sp', on_release=lambda x: self.pass_dialog.dismiss()))
        btns.add_widget(self._ui_card_btn('CONFIRMER', (0.0, 0.62, 0.25, 1), icon='check-circle-outline', size_hint_x=0.60, font_size='13.5sp', elevation=1, on_release=_confirm))
        content.add_widget(btns)

        self.pass_dialog = MDDialog(
            title='Changer le mot de passe',
            type='custom',
            content_cls=content,
            size_hint=(0.94, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.pass_dialog, top_pad=16, side_pad=14, bottom_pad=14, spacer_top=8)
        self.pass_dialog.open()

    def logout(self):
        self.logout_diag = self._ui_confirm_dialog(
            title='Fermeture',
            message="Voulez-vous vraiment fermer l'application ?",
            confirm_text='FERMER',
            icon='logout',
            confirm_icon='power',
            accent_color=(0.82, 0.12, 0.12, 1),
            soft_bg=(0.96, 0.97, 0.99, 1),
            border_color=(0.88, 0.90, 0.94, 1),
            on_confirm=lambda: self.stop(),
        )

    # --------------------------------------------------------------------------
    # 7.3 TABLEAU DE BORD & STATISTIQUES FINANCIÈRES (CALCUL ASYNCHRONE)
    # --------------------------------------------------------------------------
    def _build_dashboard_screen(self):
        screen = MDScreen(name='dashboard')
        screen.md_bg_color = (0.95, 0.96, 0.98, 1)
        layout = MDBoxLayout(orientation='vertical')

        self.dash_toolbar = MDCard(
            orientation='horizontal',
            size_hint=(1, None),
            height=dp(64),
            radius=[0, 0, 20, 20],
            elevation=3,
            md_bg_color=(0.09, 0.44, 0.84, 1),
            padding=[dp(12), dp(8), dp(12), dp(8)],
            spacing=dp(10),
        )

        btn_history = MDCard(
            orientation='horizontal',
            size_hint_y=None,
            height=dp(40),
            adaptive_width=True,
            radius=[20],
            elevation=0,
            md_bg_color=(1, 1, 1, 0.18),
            line_color=(1, 1, 1, 0.28),
            line_width=1,
            ripple_behavior=True,
            padding=[dp(14), 0, dp(16), 0],
            spacing=dp(8),
            pos_hint={'center_y': 0.5},
            on_release=lambda x: self.show_pending_dialog(),
        )
        btn_history.add_widget(MDIcon(icon='history', font_size='22sp', size_hint=(None, None), size=(dp(24), dp(24)), halign='center', valign='middle', pos_hint={'center_y': 0.5}, theme_text_color='Custom', text_color=(1, 1, 1, 1)))
        btn_history.add_widget(MDLabel(text='HISTORIQUE', bold=True, font_size='13sp', theme_text_color='Custom', text_color=(1, 1, 1, 1), adaptive_width=True, valign='middle', pos_hint={'center_y': 0.5}))
        self.dash_toolbar.add_widget(btn_history)
        self.dash_toolbar.add_widget(MDBoxLayout(size_hint_x=1))

        actions_hdr = MDBoxLayout(orientation='horizontal', adaptive_width=True, spacing=dp(8), pos_hint={'center_y': 0.5})
        actions_hdr.add_widget(self._ui_icon_pill('cog-outline', (1, 1, 1, 0.16), (1, 1, 1, 1), on_release=lambda x: self.open_settings_menu(), size_dp=40, radius=20, icon_size='21sp'))
        actions_hdr.add_widget(self._ui_icon_pill('logout', (0.95, 0.25, 0.25, 0.35), (1, 1, 1, 1), on_release=lambda x: self.logout(), size_dp=40, radius=20, icon_size='21sp'))
        self.dash_toolbar.add_widget(actions_hdr)

        layout.add_widget(self.dash_toolbar)
        scroll = MDScrollView()
        self.main_dash_content = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(18), padding=dp(15))
        self.buttons_container = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(14))
        self.main_dash_content.add_widget(self.buttons_container)

        self.stats_card_container = MDCard(
            orientation='vertical',
            size_hint_y=None,
            adaptive_height=True,
            padding=dp(15),
            radius=[20],
            elevation=1,
            md_bg_color=(1, 1, 1, 1),
        )
        self.main_dash_content.add_widget(self.stats_card_container)
        scroll.add_widget(self.main_dash_content)
        layout.add_widget(scroll)
        screen.add_widget(layout)
        return screen

    def _create_dash_btn(self, icon: str, text: str, bg_color, icon_color, action):
        card = MDCard(
            orientation='vertical',
            padding=[dp(8), dp(12), dp(8), dp(10)],
            spacing=dp(8),
            radius=[16],
            ripple_behavior=True,
            on_release=action,
            md_bg_color=bg_color,
            line_color=(0, 0, 0, 0.06),
            line_width=1,
            elevation=1,
            size_hint_y=None,
            height=dp(106),
        )
        badge = self._ui_icon_pill(icon, (1, 1, 1, 0.85), icon_color, size_dp=50, radius=25, icon_size='30sp')
        badge.pos_hint = {'center_x': 0.5}
        card.add_widget(badge)
        card.add_widget(
            MDLabel(
                text=text,
                halign='center',
                valign='middle',
                bold=True,
                font_size='12.5sp',
                theme_text_color='Custom',
                text_color=(0.15, 0.17, 0.22, 1),
                adaptive_height=True,
            )
        )
        return card

    def update_dashboard_layout(self):
        if not getattr(self, 'buttons_container', None) or not getattr(self, 'stats_card_container', None):
            return
        if getattr(self, '_updating_dashboard_ui', False):
            return
        self._updating_dashboard_ui = True
        try:
            if self.sm and self.sm.has_screen('dashboard'):
                self.sm.get_screen('dashboard').md_bg_color = (0.95, 0.96, 0.98, 1)
            self.buttons_container.clear_widgets()

            c_bv = ((0.88, 0.97, 0.9, 1), (0.08, 0.58, 0.2, 1))
            c_ba = ((1.0, 0.94, 0.84, 1), (0.88, 0.4, 0.0, 1))
            c_fc = ((0.88, 0.94, 1.0, 1), (0.0, 0.32, 0.75, 1))
            c_ff = ((0.98, 0.88, 0.88, 1), (0.72, 0.1, 0.2, 1))
            c_fp = ((0.94, 0.88, 1.0, 1), (0.5, 0.1, 0.65, 1))
            c_dp = ((0.85, 0.97, 0.97, 1), (0.0, 0.5, 0.5, 1))
            c_bs = ((0.95, 0.92, 0.88, 1), (0.5, 0.3, 0.1, 1))
            c_be = ((0.9, 0.91, 0.99, 1), (0.2, 0.25, 0.72, 1))
            c_rc = ((1.0, 0.86, 0.86, 1), (0.85, 0.05, 0.05, 1))
            c_rf = ((1.0, 0.89, 0.95, 1), (0.78, 0.1, 0.5, 1))
            c_tr = ((0.91, 0.94, 0.97, 1), (0.28, 0.38, 0.48, 1))
            c_pr = ((1.0, 0.96, 0.84, 1), (0.72, 0.55, 0.0, 1))
            c_cl = ((0.86, 0.95, 1.0, 1), (0.0, 0.55, 0.78, 1))
            c_fo = ((0.96, 0.93, 0.9, 1), (0.58, 0.38, 0.2, 1))
            c_de = ((0.99, 0.88, 0.88, 1), (0.8, 0.05, 0.05, 1))

            if self.is_seller_mode:
                self.buttons_container.add_widget(self._create_dash_btn('cart', 'VENTE (BV)', c_bv[0], c_bv[1], lambda x: self.open_mode('sale')))
                grid = MDGridLayout(cols=2, spacing=dp(10), adaptive_height=True)
                grid.add_widget(self._create_dash_btn('keyboard-return', 'RETOUR CLIENT', c_rc[0], c_rc[1], lambda x: self.open_mode('return_sale')))
                grid.add_widget(self._create_dash_btn('account-group', 'CLIENTS', c_cl[0], c_cl[1], lambda x: self.open_entity_manager(AppConstants.ENTITY_CLIENT)))
                self.buttons_container.add_widget(grid)
            else:
                grid = MDGridLayout(cols=2, spacing=dp(10), adaptive_height=True)
                grid.add_widget(self._create_dash_btn('cart', 'VENTE (BV)', c_bv[0], c_bv[1], lambda x: self.open_mode('sale')))
                grid.add_widget(self._create_dash_btn('truck', 'ACHAT (BA)', c_ba[0], c_ba[1], lambda x: self.open_mode('purchase')))
                grid.add_widget(self._create_dash_btn('file-document', 'FACTURE (FC)', c_fc[0], c_fc[1], lambda x: self.open_mode('invoice_sale')))
                grid.add_widget(self._create_dash_btn('file-document-edit', 'FACT. ACHAT (FF)', c_ff[0], c_ff[1], lambda x: self.open_mode('invoice_purchase')))
                grid.add_widget(self._create_dash_btn('file-document-outline', 'PROFORMA (FP)', c_fp[0], c_fp[1], lambda x: self.open_mode('proforma')))
                grid.add_widget(self._create_dash_btn('clipboard-list', 'COMMANDE (DP)', c_dp[0], c_dp[1], lambda x: self.open_mode('order_purchase')))
                grid.add_widget(self._create_dash_btn('tray-arrow-up', 'SORTIE (BS)', c_bs[0], c_bs[1], lambda x: self.open_mode('stock_out')))
                grid.add_widget(self._create_dash_btn('tray-arrow-down', 'ENTRÉE (BE)', c_be[0], c_be[1], lambda x: self.open_mode('stock_in')))
                grid.add_widget(self._create_dash_btn('keyboard-return', 'RETOUR CL.', c_rc[0], c_rc[1], lambda x: self.open_mode('return_sale')))
                grid.add_widget(self._create_dash_btn('undo', 'RETOUR FR.', c_rf[0], c_rf[1], lambda x: self.open_mode('return_purchase')))
                grid.add_widget(self._create_dash_btn('transfer', 'TRANSFERT (TR)', c_tr[0], c_tr[1], lambda x: self.open_mode('transfer')))
                grid.add_widget(self._create_dash_btn('database-edit', 'PRODUITS', c_pr[0], c_pr[1], lambda x: self.open_mode('manage_products')))
                grid.add_widget(self._create_dash_btn('account-group', 'CLIENTS', c_cl[0], c_cl[1], lambda x: self.open_entity_manager(AppConstants.ENTITY_CLIENT)))
                grid.add_widget(self._create_dash_btn('truck-delivery', 'FOURNISSEURS', c_fo[0], c_fo[1], lambda x: self.open_entity_manager(AppConstants.ENTITY_SUPPLIER)))
                self.buttons_container.add_widget(grid)
                self.buttons_container.add_widget(self._create_dash_btn('cash-minus', 'DÉPENSES / CHARGES', c_de[0], c_de[1], lambda x: self.open_add_expense_dialog()))

            for attr_name in (
                'lbl_stat_caisse', 'lbl_stat_sales', 'lbl_stat_in', 'lbl_stat_purchases',
                'lbl_stat_out', 'lbl_stat_profit', 'lbl_stat_expenses',
                'lbl_stat_stock_val', 'lbl_stat_client_debts', 'lbl_stat_supplier_dues',
            ):
                setattr(self, attr_name, None)

            self.stats_card_container.clear_widgets()
            self.stats_card_container.padding = dp(14)
            self.stats_card_container.spacing = dp(12)
            self.stats_card_container.line_color = (0.89, 0.91, 0.94, 1)
            self.stats_card_container.line_width = 1
            self.stats_card_container.elevation = 0

            header_box = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True, spacing=dp(10))
            self.lbl_dashboard_date = MDLabel(
                text="Tableau de Bord (Aujourd'hui)",
                font_style='Subtitle1',
                font_size='17sp',
                bold=True,
                theme_text_color='Primary',
                halign='center',
                size_hint_y=None,
                height=dp(26),
            )
            header_box.add_widget(self.lbl_dashboard_date)

            tabs_container, self.dash_period_btns = self._ui_period_tabs(
                on_filter=self.filter_dashboard_stats,
                on_calendar=self.open_stats_date_picker,
            )
            header_box.add_widget(tabs_container)
            self.stats_card_container.add_widget(header_box)

            def _stat_card(icon: str, title: str, attr_name: str, bg_c, ic_c):
                c = MDCard(
                    orientation='vertical',
                    padding=[dp(10), dp(10)],
                    spacing=dp(8),
                    radius=[14],
                    md_bg_color=bg_c,
                    line_color=(ic_c[0], ic_c[1], ic_c[2], 0.22),
                    line_width=1,
                    elevation=0,
                    size_hint_y=None,
                    height=dp(88),
                )
                top = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(28), spacing=dp(6))
                top.add_widget(self._ui_icon_pill(icon, (1, 1, 1, 0.88), ic_c, size_dp=28, radius=14, icon_size='17sp'))
                top.add_widget(
                    MDLabel(
                        text=title,
                        halign='left',
                        valign='middle',
                        font_style='Caption',
                        font_size='11.5sp',
                        bold=True,
                        theme_text_color='Custom',
                        text_color=(0.35, 0.38, 0.45, 1),
                        shorten=True,
                    )
                )
                lbl = MDLabel(
                    text='0,00 DA',
                    halign='center',
                    valign='middle',
                    font_style='Subtitle1',
                    font_size='16.5sp',
                    bold=True,
                    theme_text_color='Custom',
                    text_color=(0.1, 0.12, 0.16, 1),
                    shorten=True,
                )
                setattr(self, attr_name, lbl)
                c.add_widget(top)
                c.add_widget(lbl)
                return c

            if self.is_seller_mode:
                self.stats_more_container = None
                stats_grid = MDGridLayout(cols=2, spacing=dp(10), adaptive_height=True)
                stats_grid.add_widget(_stat_card('cart-outline', 'Ventes Nettes', 'lbl_stat_sales', (0.91, 0.95, 1, 1), (0.06, 0.4, 0.82, 1)))
                stats_grid.add_widget(_stat_card('cash-plus', 'Encaissements', 'lbl_stat_in', (0.91, 0.98, 0.92, 1), (0.04, 0.58, 0.22, 1)))
                self.stats_card_container.add_widget(stats_grid)
            else:
                caisse_card = MDCard(
                    orientation='vertical',
                    radius=[16],
                    md_bg_color=(0.94, 0.99, 0.96, 1),
                    line_color=(0.74, 0.89, 0.79, 1),
                    line_width=1.2,
                    elevation=0,
                    size_hint_y=None,
                    height=dp(108),
                    padding=[dp(14), dp(12)],
                    spacing=dp(4),
                )
                caisse_top = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(28), spacing=dp(8), adaptive_width=True, pos_hint={'center_x': 0.5})
                caisse_top.add_widget(self._ui_icon_pill('cash-multiple', (0.84, 0.96, 0.88, 1), (0.04, 0.55, 0.24, 1), size_dp=28, radius=14, icon_size='18sp'))
                caisse_top.add_widget(MDLabel(text='TOTAL EN CAISSE', bold=True, font_style='Subtitle2', theme_text_color='Custom', text_color=(0.06, 0.38, 0.18, 1), adaptive_width=True, valign='middle', pos_hint={'center_y': 0.5}))
                caisse_card.add_widget(caisse_top)

                caisse_card.add_widget(
                    MDLabel(
                        text='(Ventes payées + Versements - Dépenses)',
                        halign='center',
                        font_style='Caption',
                        font_size='11sp',
                        theme_text_color='Secondary',
                        size_hint_y=None,
                        height=dp(16),
                    )
                )
                self.lbl_stat_caisse = MDLabel(
                    text='0,00 DA',
                    halign='center',
                    valign='middle',
                    font_style='H5',
                    font_size='23sp',
                    bold=True,
                    theme_text_color='Custom',
                    text_color=(0.08, 0.1, 0.14, 1),
                )
                caisse_card.add_widget(self.lbl_stat_caisse)
                self.stats_card_container.add_widget(caisse_card)

                self.is_stats_expanded = False
                self.btn_stats_toggle = MDCard(
                    orientation='horizontal',
                    size_hint_y=None,
                    height=dp(42),
                    radius=[12],
                    md_bg_color=(0.93, 0.96, 1, 1),
                    line_color=(0.82, 0.88, 0.97, 1),
                    line_width=1,
                    elevation=0,
                    ripple_behavior=True,
                    spacing=dp(6),
                    on_release=self.toggle_admin_stats_details,
                )
                self.lbl_stats_toggle = MDLabel(
                    text='Voir plus de détails',
                    font_style='Caption',
                    bold=True,
                    theme_text_color='Custom',
                    text_color=(0.09, 0.44, 0.84, 1),
                    adaptive_width=True,
                    pos_hint={'center_y': 0.5},
                )
                self.ico_stats_toggle = MDIcon(
                    icon='chevron-down',
                    theme_text_color='Custom',
                    text_color=(0.09, 0.44, 0.84, 1),
                    font_size='20sp',
                    size_hint_x=None,
                    width=dp(22),
                    pos_hint={'center_y': 0.5},
                )
                self.btn_stats_toggle.add_widget(MDBoxLayout())
                self.btn_stats_toggle.add_widget(self.lbl_stats_toggle)
                self.btn_stats_toggle.add_widget(self.ico_stats_toggle)
                self.btn_stats_toggle.add_widget(MDBoxLayout())
                self.stats_card_container.add_widget(self.btn_stats_toggle)

                self.stats_more_container = MDBoxLayout(orientation='vertical', spacing=dp(10), adaptive_height=True)
                stats_grid = MDGridLayout(cols=2, spacing=dp(10), adaptive_height=True)
                stats_grid.add_widget(_stat_card('cart-outline', 'Ventes Nettes', 'lbl_stat_sales', (0.91, 0.95, 1, 1), (0.06, 0.4, 0.82, 1)))
                stats_grid.add_widget(_stat_card('cash-plus', 'Encaissements', 'lbl_stat_in', (0.91, 0.98, 0.92, 1), (0.04, 0.58, 0.22, 1)))
                stats_grid.add_widget(_stat_card('truck-outline', 'Achats Nets', 'lbl_stat_purchases', (1, 0.96, 0.88, 1), (0.88, 0.45, 0, 1)))
                stats_grid.add_widget(_stat_card('cash-minus', 'Décaissements', 'lbl_stat_out', (1, 0.92, 0.92, 1), (0.82, 0.12, 0.12, 1)))
                stats_grid.add_widget(_stat_card('chart-line-variant', 'Bénéfice Net', 'lbl_stat_profit', (0.95, 0.91, 1, 1), (0.48, 0.12, 0.72, 1)))
                stats_grid.add_widget(_stat_card('cash-remove', 'Dépenses', 'lbl_stat_expenses', (1, 0.94, 0.89, 1), (0.85, 0.35, 0, 1)))
                stats_grid.add_widget(_stat_card('account-cash', 'Créances Clients', 'lbl_stat_client_debts', (0.91, 0.99, 0.93, 1), (0.04, 0.58, 0.22, 1)))
                stats_grid.add_widget(_stat_card('account-credit-card', 'Dettes Fourniss.', 'lbl_stat_supplier_dues', (1, 0.91, 0.91, 1), (0.8, 0.08, 0.08, 1)))
                self.stats_more_container.add_widget(stats_grid)

                stock_card = _stat_card('package-variant', 'Stock (Coût)', 'lbl_stat_stock_val', (0.92, 0.97, 0.97, 1), (0, 0.5, 0.52, 1))
                stock_card.height = dp(82)
                self.stats_more_container.add_widget(stock_card)

            if self.main_dash_content.parent is not None:
                self.main_dash_content.parent.scroll_y = 1.0

            if getattr(self, '_dash_stats_ev', None):
                self._dash_stats_ev.cancel()
            self._dash_stats_ev = Clock.schedule_once(lambda dt: self.filter_dashboard_stats('auj'), 0.05)
        finally:
            self._updating_dashboard_ui = False

    def toggle_admin_stats_details(self, *args):
        if not getattr(self, 'stats_card_container', None) or not getattr(self, 'stats_more_container', None):
            return
        self.is_stats_expanded = not getattr(self, 'is_stats_expanded', False)
        if self.is_stats_expanded:
            if self.stats_more_container.parent is None:
                self.stats_card_container.add_widget(self.stats_more_container)
            if getattr(self, 'lbl_stats_toggle', None):
                self.lbl_stats_toggle.text = 'Masquer les détails'
            if getattr(self, 'ico_stats_toggle', None):
                self.ico_stats_toggle.icon = 'chevron-up'
        else:
            if self.stats_more_container.parent is not None:
                self.stats_card_container.remove_widget(self.stats_more_container)
            if getattr(self, 'lbl_stats_toggle', None):
                self.lbl_stats_toggle.text = 'Voir plus de détails'
            if getattr(self, 'ico_stats_toggle', None):
                self.ico_stats_toggle.icon = 'chevron-down'

    def open_stats_date_picker(self, instance):
        dp_dlg = MDDatePicker(mode='range')
        dp_dlg.bind(on_save=lambda inst, val, dr: self.filter_dashboard_stats('custom', dr[0], dr[-1]) if dr else self.filter_dashboard_stats('custom', val, val))
        dp_dlg.open()

    def filter_dashboard_stats(self, filter_type: str = 'auj', specific_start=None, specific_end=None):
        today = datetime.now().date()
        start_d, end_d = today, today
        cal_lbl = 'CAL'
        active_k = filter_type

        if filter_type == 'custom' and specific_start:
            start_d, end_d = specific_start, (specific_end or specific_start)
            cal_lbl = start_d.strftime('%d/%m') if start_d == end_d else 'CAL'
            self.lbl_dashboard_date.text = f"Tableau de Bord ({start_d.strftime('%d/%m')} - {end_d.strftime('%d/%m')})" if start_d != end_d else f"Tableau de Bord ({start_d.strftime('%d/%m/%Y')})"
        elif filter_type == 'hier':
            start_d = end_d = today - timedelta(days=1)
            self.lbl_dashboard_date.text = 'Tableau de Bord (Hier)'
        elif filter_type == 'semaine':
            start_d = today - timedelta(days=(today.weekday() + 2) % 7)
            self.lbl_dashboard_date.text = 'Tableau de Bord (Cette semaine)'
        elif filter_type == 'mois':
            start_d = today.replace(day=1)
            self.lbl_dashboard_date.text = 'Tableau de Bord (Ce mois)'
        else:
            active_k = 'auj'
            self.lbl_dashboard_date.text = "Tableau de Bord (Aujourd'hui)"

        if getattr(self, 'dash_period_btns', None):
            self._activate_period_tab(self.dash_period_btns, active_k, cal_lbl)

        self.stats_start_date = str(start_d)
        self.stats_end_date = str(end_d)
        self.check_and_load_stats()

    def check_and_load_stats(self):
        today_str = datetime.now().strftime('%Y-%m-%d')
        s_date = getattr(self, 'stats_start_date', today_str)
        e_date = getattr(self, 'stats_end_date', today_str)
        self._stats_seq += 1
        seq = self._stats_seq
        threading.Thread(target=self._stats_worker, args=(s_date, e_date, seq), daemon=True).start()

    def _stats_worker(self, s_date: str, e_date: str, seq: int):
        try:
            st = self.db.get_comprehensive_stats(s_date, e_date)
            if seq == self._stats_seq:
                self._apply_stats_ui(st, seq)
        except Exception:
            pass

    @mainthread
    def _apply_stats_ui(self, st: dict, seq: int):
        if seq != self._stats_seq:
            return
        if getattr(self, 'lbl_stat_caisse', None):
            caisse_net = st['caisse_total']
            self.lbl_stat_caisse.text = FinancialEngine.format_money(caisse_net)
            self.lbl_stat_caisse.text_color = (0.8, 0, 0, 1) if caisse_net < 0 else (0.1, 0.1, 0.1, 1)
        if getattr(self, 'lbl_stat_sales', None):
            self.lbl_stat_sales.text = FinancialEngine.format_money(st['sales'])
        if getattr(self, 'lbl_stat_in', None):
            self.lbl_stat_in.text = FinancialEngine.format_money(st['cash_in'])
        if getattr(self, 'lbl_stat_purchases', None):
            self.lbl_stat_purchases.text = FinancialEngine.format_money(st['purchases'])
        if getattr(self, 'lbl_stat_out', None):
            self.lbl_stat_out.text = FinancialEngine.format_money(st['cash_out'])
        if getattr(self, 'lbl_stat_expenses', None):
            self.lbl_stat_expenses.text = FinancialEngine.format_money(st['expenses'])
        if getattr(self, 'lbl_stat_stock_val', None):
            self.lbl_stat_stock_val.text = FinancialEngine.format_money(st['stock_value_cost'])
        if getattr(self, 'lbl_stat_client_debts', None):
            self.lbl_stat_client_debts.text = FinancialEngine.format_money(st['client_debts'])
        if getattr(self, 'lbl_stat_supplier_dues', None):
            self.lbl_stat_supplier_dues.text = FinancialEngine.format_money(st['supplier_dues'])
        if getattr(self, 'lbl_stat_profit', None):
            prof = st['profit']
            self.lbl_stat_profit.text = FinancialEngine.format_money(prof)
            self.lbl_stat_profit.text_color = (0.8, 0, 0, 1) if prof < 0 else (0, 0.55, 0.35, 1)

    # --------------------------------------------------------------------------
    # 7.4 ÉCRAN PRODUITS, RECHERCHE, FAMILLES & FICHE PRODUIT
    # --------------------------------------------------------------------------
    def _create_modern_header(self, title: str, on_back, mode_icon: str = 'cart', on_right=None):
        class _ModernHeaderBar(MDCard):
            title = StringProperty('')
            subtitle = StringProperty('')
            mode_icon = StringProperty('cart')

            def __init__(self, back_cb, right_cb, **kwargs):
                super().__init__(**kwargs)
                self.orientation = 'horizontal'
                self.size_hint = (1, None)
                self.height = dp(64)
                self.radius = [0, 0, 20, 20]
                self.elevation = 3
                self.padding = [dp(12), dp(8), dp(12), dp(8)]
                self.spacing = dp(12)

                btn_back = MDCard(
                    size_hint=(None, None),
                    size=(dp(40), dp(40)),
                    radius=[20],
                    elevation=0,
                    md_bg_color=(1, 1, 1, 0.18),
                    line_color=(1, 1, 1, 0.25),
                    line_width=1,
                    ripple_behavior=True,
                    pos_hint={'center_y': 0.5},
                    on_release=back_cb,
                )
                back_fl = MDFloatLayout()
                back_fl.add_widget(
                    MDIcon(
                        icon='arrow-left',
                        font_size='22sp',
                        size_hint=(None, None),
                        size=(dp(24), dp(24)),
                        halign='center',
                        valign='middle',
                        pos_hint={'center_x': 0.5, 'center_y': 0.5},
                        theme_text_color='Custom',
                        text_color=(1, 1, 1, 1),
                    )
                )
                btn_back.add_widget(back_fl)
                self.add_widget(btn_back)

                center_box = MDBoxLayout(
                    orientation='vertical',
                    size_hint_x=1,
                    adaptive_height=True,
                    spacing=dp(2),
                    pos_hint={'center_y': 0.5},
                )
                self._lbl_title = MDLabel(
                    text=self.title,
                    bold=True,
                    font_size='18sp',
                    theme_text_color='Custom',
                    text_color=(1, 1, 1, 1),
                    halign='left',
                    shorten=True,
                    shorten_from='right',
                    size_hint_y=None,
                    height=dp(24),
                )
                self._lbl_sub = MDLabel(
                    text=self.subtitle,
                    bold=True,
                    font_size='12sp',
                    font_name='ArabicFont',
                    theme_text_color='Custom',
                    text_color=(1, 1, 1, 0.9),
                    halign='left',
                    shorten=True,
                    shorten_from='right',
                    size_hint_y=None,
                    height=0,
                    opacity=0,
                )
                center_box.add_widget(self._lbl_title)
                center_box.add_widget(self._lbl_sub)
                self.add_widget(center_box)

                btn_right = MDCard(
                    size_hint=(None, None),
                    size=(dp(40), dp(40)),
                    radius=[20],
                    elevation=0,
                    md_bg_color=(1, 1, 1, 0.18),
                    line_color=(1, 1, 1, 0.25 if right_cb else 0.12),
                    line_width=1,
                    ripple_behavior=bool(right_cb),
                    pos_hint={'center_y': 0.5},
                )
                if right_cb:
                    btn_right.bind(on_release=right_cb)
                right_fl = MDFloatLayout()
                self._right_ico = MDIcon(
                    icon=self.mode_icon,
                    font_size='21sp',
                    size_hint=(None, None),
                    size=(dp(24), dp(24)),
                    halign='center',
                    valign='middle',
                    pos_hint={'center_x': 0.5, 'center_y': 0.5},
                    theme_text_color='Custom',
                    text_color=(1, 1, 1, 1),
                )
                right_fl.add_widget(self._right_ico)
                btn_right.add_widget(right_fl)
                self.add_widget(btn_right)

                def _on_sub(inst, val):
                    self._lbl_sub.text = val
                    has_s = bool(str(val or '').strip())
                    self._lbl_sub.height = dp(16) if has_s else 0
                    self._lbl_sub.opacity = 1 if has_s else 0

                self.bind(
                    title=lambda inst, val: setattr(self._lbl_title, 'text', val),
                    subtitle=_on_sub,
                    mode_icon=lambda inst, val: setattr(self._right_ico, 'icon', val),
                )

        return _ModernHeaderBar(
            back_cb=on_back,
            right_cb=on_right,
            title=title,
            mode_icon=mode_icon,
            md_bg_color=(0.09, 0.44, 0.84, 1),
        )

    def _on_prod_header_right_tap(self, instance=None):
        if self.current_mode == 'transfer':
            self.toggle_location()
        elif self.current_mode == 'manage_products':
            self.open_filter_menu(None)
        elif self.current_mode != 'initial_stock':
            self.show_entity_selection_dialog(None)

    def _build_products_screen(self):
        screen = MDScreen(name='products')
        screen.md_bg_color = (0.95, 0.96, 0.98, 1)
        root = MDBoxLayout(orientation='vertical')
        header = MDBoxLayout(orientation='vertical', adaptive_height=True, md_bg_color=(1, 1, 1, 1))
        self.prod_toolbar = self._create_modern_header(
            'Produits',
            on_back=lambda x: self.go_back(),
            mode_icon='cart',
            on_right=self._on_prod_header_right_tap,
        )
        header.add_widget(self.prod_toolbar)

        self.prod_search_layout = MDBoxLayout(
            orientation='horizontal',
            padding=[dp(12), dp(10), dp(12), dp(6)],
            spacing=dp(8),
            size_hint_y=None,
            height=dp(62),
            md_bg_color=(1, 1, 1, 1),
        )
        search_card = MDCard(
            orientation='horizontal',
            size_hint_x=1,
            size_hint_y=None,
            height=dp(46),
            radius=[12],
            elevation=0,
            md_bg_color=(0.95, 0.96, 0.98, 1),
            line_color=(0.86, 0.88, 0.93, 1),
            line_width=1,
            padding=[dp(12), 0, dp(8), 0],
            spacing=dp(6),
            pos_hint={'center_y': 0.5},
        )
        search_card.add_widget(
            MDIcon(
                icon='magnify',
                theme_text_color='Secondary',
                font_size='20sp',
                size_hint_x=None,
                width=dp(22),
                pos_hint={'center_y': 0.5},
            )
        )
        self.search_field = SmartTextField(
            hint_text='Rechercher...',
            mode='line',
            line_color_normal=(0, 0, 0, 0),
            line_color_focus=(0, 0, 0, 0),
            pos_hint={'center_y': 0.5},
        )
        self.search_field.bind(text=self.filter_products)
        search_card.add_widget(self.search_field)
        self.prod_search_layout.add_widget(search_card)

        self.btn_scan_prod = self._ui_icon_pill(
            'barcode-scan',
            (0.93, 0.95, 0.98, 1),
            (0.15, 0.25, 0.4, 1),
            on_release=self.open_barcode_scanner,
            size_dp=46,
            radius=12,
            icon_size='24sp',
            border_color=(0.86, 0.88, 0.93, 1),
        )
        self.prod_search_layout.add_widget(self.btn_scan_prod)
        self.btn_add_prod = self._ui_icon_pill(
            'plus',
            (0.0, 0.65, 0.25, 1),
            (1, 1, 1, 1),
            on_release=lambda x: self.show_manage_product_dialog(None),
            size_dp=46,
            radius=12,
            icon_size='26sp',
        )
        self.prod_search_layout.add_widget(self.btn_add_prod)
        header.add_widget(self.prod_search_layout)

        filter_box = MDBoxLayout(orientation='vertical', size_hint_y=None, height=dp(52), padding=[dp(12), 0, dp(12), dp(8)])
        family_card = MDCard(
            size_hint=(1, 1),
            radius=[12],
            md_bg_color=(0.95, 0.96, 0.98, 1),
            line_color=(0.86, 0.88, 0.93, 1),
            line_width=1,
            elevation=0,
            ripple_behavior=True,
            on_release=self.open_filter_menu,
        )
        card_c = MDBoxLayout(orientation='horizontal', padding=[dp(12), 0, dp(10), 0], spacing=dp(8))
        card_c.add_widget(
            MDIcon(
                icon='shape-outline',
                theme_text_color='Primary',
                font_size='19sp',
                size_hint_x=None,
                width=dp(22),
                pos_hint={'center_y': 0.5},
            )
        )
        self.btn_main_family_filter = MDLabel(
            text=AppConstants.FAM_ALL,
            theme_text_color='Custom',
            text_color=(0.15, 0.17, 0.22, 1),
            font_style='Subtitle2',
            bold=True,
            halign='left',
            valign='center',
            font_name='ArabicFont',
        )
        card_c.add_widget(self.btn_main_family_filter)
        card_c.add_widget(
            MDIcon(
                icon='chevron-down',
                pos_hint={'center_y': 0.5},
                theme_text_color='Secondary',
                size_hint_x=None,
                width=dp(22),
            )
        )
        family_card.add_widget(card_c)
        filter_box.add_widget(family_card)
        header.add_widget(filter_box)
        header.add_widget(MDSeparator(height=dp(1), color=(0.89, 0.91, 0.94, 1)))
        root.add_widget(header)

        self.rv_products = ProductRecycleView()
        self.rv_products.size_hint_y = 1
        root.add_widget(self.rv_products)

        self.cart_bar = MDCard(
            size_hint_y=None,
            height=dp(60),
            padding=[dp(16), dp(8)],
            md_bg_color=(0.09, 0.44, 0.84, 1),
            radius=[18, 18, 0, 0],
            ripple_behavior=True,
            on_release=self.open_cart_screen,
            elevation=2,
        )
        cart_box = MDBoxLayout(orientation='horizontal', spacing=dp(10))
        cart_box.add_widget(
            MDIcon(
                icon='cart-outline',
                theme_text_color='Custom',
                text_color=(1, 1, 1, 1),
                font_size='23sp',
                size_hint_x=None,
                width=dp(25),
                pos_hint={'center_y': 0.5},
            )
        )
        self.lbl_cart_count = MDLabel(
            text='PANIER (0)',
            theme_text_color='Custom',
            text_color=(1, 1, 1, 1),
            bold=True,
            halign='left',
            valign='middle',
            size_hint_x=0.48,
            font_style='Subtitle1',
        )
        self.lbl_cart_total = MDLabel(
            text='0,00 DA',
            theme_text_color='Custom',
            text_color=(1, 1, 1, 1),
            bold=True,
            halign='right',
            valign='middle',
            font_style='H6',
            size_hint_x=0.52,
        )
        cart_box.add_widget(self.lbl_cart_count)
        cart_box.add_widget(self.lbl_cart_total)
        cart_box.add_widget(
            MDIcon(
                icon='chevron-right',
                theme_text_color='Custom',
                text_color=(1, 1, 1, 0.9),
                font_size='23sp',
                size_hint_x=None,
                width=dp(22),
                pos_hint={'center_y': 0.5},
            )
        )
        self.cart_bar.add_widget(cart_box)
        root.add_widget(self.cart_bar)
        screen.add_widget(root)
        return screen

    def open_mode(self, mode: str, skip_dialog: bool = False):
        self.current_mode = mode
        if not skip_dialog:
            self.cart = []
            self.selected_entity = None
            self.editing_tx_id = None
            self.editing_doc_type = None
            self.editing_payment_amount = Decimal('0.00')
            self.editing_payment_method = AppConstants.PAY_DEFAULT
            self.editing_global_discount = Decimal('0.00')
        self.selected_location = AppConstants.LOC_STORE
        self.update_cart_button()
        self.selected_family_filter = AppConstants.FAM_ALL
        if hasattr(self, 'btn_main_family_filter'):
            self.btn_main_family_filter.text = AppConstants.FAM_ALL

        mode_styles = {
            'sale': ((0.06, 0.56, 0.24, 1), 'account-switch-outline'),
            'purchase': ((0.88, 0.42, 0.0, 1), 'account-switch-outline'),
            'invoice_sale': ((0.09, 0.44, 0.84, 1), 'account-switch-outline'),
            'invoice_purchase': ((0.76, 0.14, 0.22, 1), 'account-switch-outline'),
            'proforma': ((0.5, 0.12, 0.65, 1), 'account-switch-outline'),
            'order_purchase': ((0.0, 0.5, 0.52, 1), 'account-switch-outline'),
            'return_sale': ((0.82, 0.12, 0.12, 1), 'account-switch-outline'),
            'return_purchase': ((0.0, 0.55, 0.55, 1), 'account-switch-outline'),
            'stock_in': ((0.15, 0.38, 0.75, 1), 'account-switch-outline'),
            'stock_out': ((0.48, 0.28, 0.12, 1), 'account-switch-outline'),
            'transfer': ((0.5, 0.1, 0.58, 1), 'swap-horizontal'),
            'initial_stock': ((0.0, 0.5, 0.5, 1), 'database-plus'),
            'manage_products': ((0.09, 0.44, 0.84, 1), 'shape-outline'),
        }
        hdr_col, hdr_ico = mode_styles.get(mode, ((0.09, 0.44, 0.84, 1), 'cart'))

        self.prod_toolbar.title = AppConstants.MODE_TITLES.get(mode, 'Produits')
        self.prod_toolbar.md_bg_color = hdr_col
        self.prod_toolbar.mode_icon = hdr_ico
        if mode == 'manage_products':
            self.prod_toolbar.subtitle = 'Catalogue & Stock'
        elif mode == 'initial_stock':
            self.prod_toolbar.subtitle = 'Édition du Bon Initial'
        elif mode == 'transfer':
            src = AppConstants.LOC_LABELS[self.selected_location]
            dst = AppConstants.LOC_LABELS[AppConstants.LOC_WAREHOUSE if self.selected_location == AppConstants.LOC_STORE else AppConstants.LOC_STORE]
            self.prod_toolbar.subtitle = f'{src}  >>>  {dst}'
        elif self.selected_entity:
            ent_lbl = 'Fournisseur' if AppConstants.get_entity_type(mode) == AppConstants.ENTITY_SUPPLIER else 'Client'
            self.prod_toolbar.subtitle = f"{ent_lbl} : {self.fix_text(self.selected_entity.get('name', ''))}"
        else:
            self.prod_toolbar.subtitle = ''

        self.theme_cls.primary_palette = AppConstants.MODE_COLORS.get(mode, 'Blue')
        self.cart_bar.md_bg_color = hdr_col

        if mode == 'manage_products':
            if self.btn_add_prod not in self.prod_search_layout.children:
                self.prod_search_layout.add_widget(self.btn_add_prod)
            if self.btn_scan_prod in self.prod_search_layout.children:
                self.prod_search_layout.remove_widget(self.btn_scan_prod)
            self.cart_bar.height = 0
            self.cart_bar.opacity = 0
            self.cart_bar.disabled = True
        else:
            if self.btn_add_prod in self.prod_search_layout.children:
                self.prod_search_layout.remove_widget(self.btn_add_prod)
            if self.btn_scan_prod not in self.prod_search_layout.children:
                self.prod_search_layout.add_widget(self.btn_scan_prod)
            self.cart_bar.height = dp(60)
            self.cart_bar.opacity = 1
            self.cart_bar.disabled = False

        if mode != 'initial_stock':
            self.load_more_products(reset=True)

        def _enter():
            self.sm.current = 'products'
            if hasattr(self, 'search_field') and self.search_field:
                self.search_field.set_raw_value('')
                self.search_field.focus = False

        modes_needing_entity = (
            'sale', 'purchase', 'return_sale', 'return_purchase',
            'invoice_sale', 'invoice_purchase', 'proforma', 'order_purchase',
            'stock_in', 'stock_out',
        )
        if mode in modes_needing_entity and not skip_dialog:
            self.show_entity_selection_dialog(None, next_action=_enter)
        elif mode != 'initial_stock':
            _enter()

    def open_filter_menu(self, instance):
        families = self.db.get_families()
        full_list = [AppConstants.FAM_ALL] + [f for f in families if f != AppConstants.FAM_ALL]
        current_fam = getattr(self, 'selected_family_filter', AppConstants.FAM_ALL)
        options = [
            {
                'label': self.fix_text(fam),
                'code': fam,
                'icon': 'shape-outline' if fam == AppConstants.FAM_ALL else 'tag-outline',
                'color': (0.45, 0.48, 0.55, 1),
            }
            for fam in full_list
        ]
        self.filter_dialog = self._ui_filter_menu('Filtrer par famille', options, current_fam, self.apply_filter)

    def apply_filter(self, family_name: str):
        self.selected_family_filter = family_name
        self.btn_main_family_filter.text = self.fix_text(family_name)
        self.load_more_products(reset=True)

    def filter_products(self, instance, text):
        if self._search_event:
            self._search_event.cancel()
        self._search_event = Clock.schedule_once(lambda dt: self.load_more_products(reset=True), 0.25)

    def load_more_products(self, reset: bool = False):
        if self.is_loading_more and not reset:
            return
        if not reset and not self.products_has_more:
            if self.rv_products:
                self.rv_products.loading_lock = False
            return
        if reset:
            self.current_page_offset = 0
            self.products_has_more = True
            self.is_loading_more = False
            if self.rv_products:
                self.rv_products.scroll_y = 1.0
                self.rv_products.data = []
                self.rv_products.refresh_from_data()
        self.is_loading_more = True
        self._prod_fetch_seq = getattr(self, '_prod_fetch_seq', 0) + 1
        seq = self._prod_fetch_seq
        threading.Thread(target=self._load_products_worker, args=(reset, seq), daemon=True).start()

    def _load_products_worker(self, reset: bool, seq: int):
        try:
            q = self.search_field.get_value().strip() if hasattr(self, 'search_field') and self.search_field else ''
            fam = getattr(self, 'selected_family_filter', AppConstants.FAM_ALL)
            doc_type = AppConstants.MODE_TO_DOC.get(self.current_mode, AppConstants.DOC_BV)
            allow_virtual = doc_type in (AppConstants.DOC_BV, AppConstants.DOC_FC, AppConstants.DOC_FP, AppConstants.DOC_DP, AppConstants.DOC_RC)
            include_def = allow_virtual and self.current_page_offset == 0 and not q and fam == AppConstants.FAM_ALL

            rows = self.db.get_products(
                limit=self.batch_size,
                offset=self.current_page_offset,
                search_query=q,
                family_filter=fam,
                include_default=include_def,
            )
            if seq != getattr(self, '_prod_fetch_seq', 0):
                return

            is_purchase = doc_type in AppConstants.PURCHASE_DOC_TYPES
            is_transfer = doc_type == AppConstants.DOC_TR
            is_qty_only = doc_type in AppConstants.QTY_ONLY_DOC_TYPES
            show_promo = doc_type in (AppConstants.DOC_BV, AppConstants.DOC_FC, AppConstants.DOC_FP) or self.current_mode == 'manage_products'
            img_dir = os.path.join(self.user_data_dir, 'product_images')
            today_d = datetime.now().date()

            rv_data = []
            for p in rows:
                is_virt = AppConstants.is_default_product(p)
                s_store = FinancialEngine.to_decimal(p.get('stock_store', 0))
                s_wh = FinancialEngine.to_decimal(p.get('stock_warehouse', 0))
                is_unlim = AppConstants.is_unlimited_stock(s_store) or AppConstants.is_unlimited_stock(s_wh)
                if (is_transfer or is_purchase) and (is_virt or is_unlim):
                    continue

                raw_img = str(p.get('image_path', '') or '')
                final_img = ''
                if raw_img:
                    if os.path.exists(raw_img):
                        final_img = raw_img
                    else:
                        cand = os.path.join(img_dir, os.path.basename(raw_img))
                        if os.path.exists(cand):
                            final_img = cand

                icon_name = 'package-variant'
                icon_col = [0, 0.6, 0.2, 1]
                thumb_bg = [0.93, 0.96, 1, 1]
                stock_col = [0.08, 0.45, 0.2, 1]
                stock_bg = [0.91, 0.97, 0.92, 1]

                if is_virt:
                    price_fmt = 'Prix Libre'
                    price_col = [0, 0.4, 0.8, 1]
                    icon_name = 'pencil-plus'
                    icon_col = [0, 0.4, 0.8, 1]
                    thumb_bg = [0.91, 0.95, 1, 1]
                    stock_txt = 'Illimité'
                    stock_col = [0.08, 0.42, 0.78, 1]
                    stock_bg = [0.91, 0.95, 1, 1]
                elif is_qty_only:
                    price_fmt = f'Tot : {FinancialEngine.format_qty(s_store + s_wh)}'
                    price_col = [0.2, 0.2, 0.8, 1]
                    stock_txt = 'Illimité' if is_unlim else f'Mag : {FinancialEngine.format_qty(s_store)} | Dép : {FinancialEngine.format_qty(s_wh)}'
                    stock_col = [0.2, 0.25, 0.32, 1]
                    stock_bg = [0.94, 0.95, 0.97, 1]
                elif is_purchase:
                    p_val = FinancialEngine.to_decimal(p.get('purchase_price', 0))
                    price_fmt = FinancialEngine.format_money(p_val)
                    price_col = [0.9, 0.5, 0, 1]
                    thumb_bg = [1, 0.96, 0.9, 1]
                    icon_col = [0.9, 0.5, 0, 1]
                    stock_txt = 'Illimité' if is_unlim else (f'Mag : {FinancialEngine.format_qty(s_store)} | Dép : {FinancialEngine.format_qty(s_wh)}' if s_wh != 0 else f'Qté : {FinancialEngine.format_qty(s_store)}')
                    if not is_unlim and (s_store + s_wh) <= 0:
                        stock_col = [0.8, 0.1, 0.1, 1]
                        stock_bg = [1, 0.92, 0.92, 1]
                else:
                    disp_price, has_promo = self._compute_product_sale_price(p, check_promo=show_promo, today_date=today_d)
                    price_fmt = FinancialEngine.format_money(disp_price)
                    if has_promo:
                        price_col = [0, 0.2, 0.8, 1]
                        icon_name = 'sale'
                        icon_col = [0.9, 0.1, 0.1, 1]
                        thumb_bg = [1, 0.92, 0.92, 1]
                    else:
                        price_col = [0, 0.6, 0.2, 1]
                    if is_unlim:
                        stock_txt = 'Illimité'
                        stock_col = [0.08, 0.42, 0.78, 1]
                        stock_bg = [0.91, 0.95, 1, 1]
                    elif s_wh != 0:
                        stock_txt = f'Qté : {FinancialEngine.format_qty(s_store)} | Dép : {FinancialEngine.format_qty(s_wh)}'
                    else:
                        stock_txt = f'Qté : {FinancialEngine.format_qty(s_store)}'
                    if not is_unlim and (s_store + s_wh) <= 0:
                        stock_col = [0.8, 0.1, 0.1, 1]
                        stock_bg = [1, 0.92, 0.92, 1]
                        if not has_promo:
                            icon_name = 'package-variant-closed'
                            icon_col = [0.8, 0, 0, 1]
                            thumb_bg = [1, 0.93, 0.93, 1]

                rv_data.append({
                    'text_name': self.fix_text(p.get('name', '')),
                    'text_price': price_fmt,
                    'text_stock': stock_txt,
                    'icon_name': icon_name,
                    'icon_color': icon_col,
                    'price_color': price_col,
                    'stock_color': stock_col,
                    'stock_bg': stock_bg,
                    'thumb_bg': thumb_bg,
                    'image_path': final_img,
                    'raw_data': p,
                })

            Clock.schedule_once(lambda dt: self._append_products_to_rv(rv_data, reset, len(rows), seq), 0)
        except Exception:
            Clock.schedule_once(lambda dt: self._finish_products_loading(), 0)

    def _compute_product_sale_price(self, product: dict, check_promo: bool = True, today_date=None) -> tuple:
        base = FinancialEngine.to_decimal(product.get('price', 0))
        cat = AppConstants.normalize_price_category(self.selected_entity.get('category', '') if self.selected_entity else '')
        final_p = base
        if cat == AppConstants.PRICE_CAT_GROS:
            w = FinancialEngine.to_decimal(product.get('price_wholesale', 0))
            if w > 0:
                final_p = w
        elif cat == AppConstants.PRICE_CAT_SEMI:
            s = FinancialEngine.to_decimal(product.get('price_semi', 0))
            if s > 0:
                final_p = s

        if check_promo and int(product.get('is_promo_active', 0) or 0) == 1:
            exp_str = str(product.get('promo_expiry', '') or '').strip()
            valid_date = True
            if exp_str:
                try:
                    exp_d = datetime.strptime(exp_str[:10], '%Y-%m-%d').date()
                    if (today_date or datetime.now().date()) > exp_d:
                        valid_date = False
                except Exception:
                    pass
            if valid_date:
                p_val = FinancialEngine.to_decimal(product.get('promo_value', 0))
                if product.get('promo_type') == 'fixed':
                    if p_val > 0:
                        return FinancialEngine.quantize_money(p_val), True
                else:
                    return FinancialEngine.quantize_money(base * (Decimal('1') - p_val / Decimal('100'))), True
        return FinancialEngine.quantize_money(final_p), False

    def _resolve_product_unit_price(self, product: dict, doc_type: str) -> tuple:
        if doc_type in AppConstants.QTY_ONLY_DOC_TYPES:
            return Decimal('0.00'), False, False

        if doc_type in (AppConstants.DOC_RC, AppConstants.DOC_RF) and self.selected_entity and not AppConstants.is_default_product(product):
            hist_price = self.db.get_last_product_price_for_entity(
                product.get('id'),
                self.selected_entity.get('id'),
                doc_type,
            )
            if hist_price is not None:
                return hist_price, False, True

        if doc_type in AppConstants.PURCHASE_DOC_TYPES:
            return FinancialEngine.quantize_money(product.get('purchase_price', 0)), False, False

        show_promo = doc_type in (AppConstants.DOC_BV, AppConstants.DOC_FC, AppConstants.DOC_FP)
        price_val, has_promo = self._compute_product_sale_price(product, check_promo=show_promo)
        return price_val, has_promo, False

    @mainthread
    def _append_products_to_rv(self, new_data: list, reset: bool, db_count: int, seq: int):
        if seq != getattr(self, '_prod_fetch_seq', 0):
            return
        self.products_has_more = db_count == self.batch_size
        if self.rv_products:
            if reset:
                self.current_page_offset = db_count
            else:
                self.current_page_offset += db_count
            self._preserve_rv_scroll_on_append(self.rv_products, new_data, reset)
        self.is_loading_more = False

    def _finish_products_loading(self):
        self.is_loading_more = False
        if self.rv_products:
            self.rv_products.loading_lock = False

    def show_zoomed_image(self, image_path: str, title: str = 'Image'):
        if not image_path or not os.path.exists(image_path):
            self.notify('Image non disponible', 'error')
            return
        content = MDBoxLayout(
            orientation='vertical',
            padding=[0, dp(4), 0, 0],
            spacing=dp(10),
            size_hint_y=None,
            height=min(dp(420), Window.height * 0.65),
        )
        img_card = MDCard(
            size_hint=(1, 1),
            radius=[14],
            elevation=0,
            line_color=(0.88, 0.90, 0.94, 1),
            line_width=1,
            padding=dp(4),
        )
        img_card.add_widget(FitImage(source=image_path, radius=[12], size_hint=(1, 1), mipmap=True))
        content.add_widget(img_card)
        content.add_widget(self._ui_card_btn('FERMER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), height=46, on_release=lambda x: self.zoom_dialog.dismiss()))

        self.zoom_dialog = MDDialog(
            title=title,
            type='custom',
            content_cls=content,
            size_hint=(0.92, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.zoom_dialog)
        self.zoom_dialog.open()

    # --- Sélecteur d'images unifié (Produits & Logo Magasin) ---
    def _open_image_picker(self, on_selected):
        self._pending_image_callback = on_selected
        if platform == 'android':
            self.open_android_image_picker()
        else:
            default_path = os.path.join(os.path.expanduser('~'), 'Pictures')
            if not os.path.exists(default_path):
                default_path = os.path.expanduser('~')
            content = MDBoxLayout(orientation='vertical', spacing=10, size_hint_y=None, height=dp(520))
            list_bg = MDCard(orientation='vertical', size_hint_y=1, md_bg_color=(0.15, 0.15, 0.15, 1), radius=[10], padding='5dp')
            self.img_file_chooser = FileChooserListView(path=default_path, filters=['*.png', '*.jpg', '*.jpeg', '*.webp'], size_hint_y=1)
            list_bg.add_widget(self.img_file_chooser)
            content.add_widget(list_bg)
            btns = MDBoxLayout(spacing=10, size_hint_y=None, height=dp(50))
            btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.4, on_release=lambda x: self.img_picker_dialog.dismiss()))
            btns.add_widget(self._ui_card_btn('SÉLECTIONNER', (0.09, 0.44, 0.84, 1), size_hint_x=0.6, on_release=lambda x: self._dispatch_image_selection(self.img_file_chooser.selection)))
            content.add_widget(btns)
            self.img_picker_dialog = MDDialog(title='Choisir une image', type='custom', content_cls=content, size_hint=(0.95, None), radius=[18, 18, 18, 18])
            self._fit_dialog(self.img_picker_dialog)
            self.img_picker_dialog.open()

    def open_image_selector(self, instance):
        self._open_image_picker(self.confirm_image_selection)

    def open_android_image_picker(self):
        try:
            from android import activity
            from jnius import autoclass, cast

            intent_cls = autoclass('android.content.Intent')
            intent = intent_cls(intent_cls.ACTION_GET_CONTENT)
            intent.setType('image/*')
            intent.addCategory(intent_cls.CATEGORY_OPENABLE)
            activity.bind(on_activity_result=self._on_image_file_chosen)
            cur_act = cast('android.app.Activity', PythonActivity.mActivity)
            cur_act.startActivityForResult(intent, 103)
        except Exception as exc:
            self.notify(f'Erreur sélecteur image : {exc}', 'error')

    def _on_image_file_chosen(self, request_code, result_code, intent):
        from android import activity
        activity.unbind(on_activity_result=self._on_image_file_chosen)
        if request_code == 103 and result_code == -1 and intent:
            uri = intent.getData()
            if uri:
                threading.Thread(target=self._copy_image_uri_worker, args=(uri,), daemon=True).start()

    def _copy_image_uri_worker(self, uri):
        try:
            from jnius import cast
            cur_act = cast('android.app.Activity', PythonActivity.mActivity)
            stream = cur_act.getContentResolver().openInputStream(uri)
            temp_path = os.path.join(self.user_data_dir, 'temp_selected_image.jpg')
            with open(temp_path, 'wb') as out:
                buf = bytearray(8192)
                while True:
                    n = stream.read(buf)
                    if n == -1:
                        break
                    out.write(buf[:n])
            stream.close()
            if os.path.exists(temp_path) and os.path.getsize(temp_path) > 0:
                Clock.schedule_once(lambda dt: self._dispatch_image_selection([temp_path]), 0)
        except Exception as exc:
            Clock.schedule_once(lambda dt, e=exc: self.notify(f'Erreur image : {e}', 'error'), 0)

    def _dispatch_image_selection(self, selection):
        if getattr(self, 'img_picker_dialog', None):
            self.img_picker_dialog.dismiss()
        cb = getattr(self, '_pending_image_callback', None)
        self._pending_image_callback = None
        if cb and selection:
            cb(selection)

    def confirm_image_selection(self, selection):
        if selection:
            self.temp_selected_image_path = selection[0]
            self.remove_image_order = False
            if hasattr(self, 'lbl_image_status'):
                self.lbl_image_status.text = f'Image : {os.path.basename(selection[0])}'
                self.lbl_image_status.text_color = (0, 0.6, 0.2, 1)
            self.notify('Image sélectionnée', 'success')

    def save_product_image_local(self, source_path: str) -> str:
        if not source_path or not os.path.exists(source_path):
            return ''
        img_dir = os.path.join(self.user_data_dir, 'product_images')
        os.makedirs(img_dir, exist_ok=True)
        dest_path = os.path.join(img_dir, f'prod_{int(time.time())}_{random.randint(1000, 9999)}.jpg')
        try:
            with Image.open(source_path) as img:
                if img.mode in ('RGBA', 'P', 'LA'):
                    bg = Image.new('RGB', img.size, (255, 255, 255))
                    rgba = img.convert('RGBA')
                    bg.paste(rgba, mask=rgba.split()[3])
                    img = bg
                elif img.mode != 'RGB':
                    img = img.convert('RGB')
                img.thumbnail((1024, 1024), Image.Resampling.LANCZOS)
                img.save(dest_path, quality=85, optimize=True)
            return dest_path
        except Exception:
            shutil.copyfile(source_path, dest_path)
            return dest_path

    def _save_store_logo_local(self, source_path: str) -> str:
        if not source_path or not os.path.exists(source_path):
            return ''
        assets_dir = os.path.join(self.user_data_dir, 'store_assets')
        os.makedirs(assets_dir, exist_ok=True)
        dest_path = os.path.join(assets_dir, f'store_logo_{int(time.time())}.png')
        try:
            with Image.open(source_path) as img:
                if img.mode in ('RGBA', 'P', 'LA'):
                    bg = Image.new('RGB', img.size, (255, 255, 255))
                    rgba = img.convert('RGBA')
                    bg.paste(rgba, mask=rgba.split()[3])
                    img = bg
                elif img.mode != 'RGB':
                    img = img.convert('RGB')
                img.thumbnail((600, 400), Image.Resampling.LANCZOS)
                img.save(dest_path, format='PNG', optimize=True)
            return dest_path
        except Exception:
            shutil.copyfile(source_path, dest_path)
            return dest_path

    def _resolve_store_logo_path(self, store_info: dict | None = None) -> str:
        si = store_info if store_info is not None else self.db.get_store_info()
        raw = str(si.get('store_logo', '') or '').strip()
        if not raw:
            return ''
        if os.path.exists(raw):
            return raw
        cand = os.path.join(self.user_data_dir, 'store_assets', os.path.basename(raw))
        if os.path.exists(cand):
            return cand
        return ''

    # --- Familles (Dialogues) ---
    def open_family_selector_dialog(self):
        families = [f for f in self.db.get_families() if f != AppConstants.FAM_ALL]
        full_list = [AppConstants.FAM_ALL] + families
        current_fam = getattr(self, 'selected_product_family_raw', AppConstants.FAM_ALL)
        options = [
            {
                'label': self.fix_text(fam),
                'code': fam,
                'icon': 'shape-outline' if fam == AppConstants.FAM_ALL else 'tag-outline',
                'color': (0.45, 0.48, 0.55, 1),
                'deletable': fam != AppConstants.FAM_ALL,
            }
            for fam in full_list
        ]
        self.family_dialog = self._ui_filter_menu(
            'Sélectionner une famille',
            options,
            current_fam,
            self._select_family_and_close,
            on_delete_item=self.show_delete_family_confirmation,
        )

    def _select_family_and_close(self, family_name: str):
        self.selected_product_family_raw = family_name
        if hasattr(self, 'btn_select_family'):
            self.btn_select_family.text = self.fix_text(family_name)

    def show_add_family_dialog(self):
        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(12),
            size_hint_y=None,
            adaptive_height=True,
            padding=[dp(2), dp(6), dp(2), dp(2)],
        )
        input_card = MDCard(
            orientation='vertical',
            size_hint=(1, None),
            height=dp(66),
            radius=[12],
            elevation=0,
            md_bg_color=(0.96, 0.97, 0.99, 1),
            line_color=(0.86, 0.88, 0.93, 1),
            line_width=1,
            padding=[dp(12), dp(4)],
        )
        field = SmartTextField(hint_text='Nom de la nouvelle famille', icon_right='tag-plus-outline', pos_hint={'center_y': 0.5})
        input_card.add_widget(field)
        content.add_widget(input_card)

        def _save(x):
            name = field.get_value().strip()
            if not name:
                self.notify('Le nom de la famille est vide.', 'error')
                return
            if self.db.add_family(name):
                self.selected_product_family_raw = name
                if hasattr(self, 'btn_select_family'):
                    self.btn_select_family.text = self.fix_text(name)
                self.add_fam_dialog.dismiss()
                self.notify(f'Famille "{name}" ajoutée.', 'success')
            else:
                self.notify('Cette famille existe déjà.', 'error')

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(46))
        btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.38, height=46, on_release=lambda x: self.add_fam_dialog.dismiss()))
        btns.add_widget(self._ui_card_btn('AJOUTER', (0.0, 0.62, 0.25, 1), icon='plus-circle-outline', size_hint_x=0.62, height=46, elevation=1, on_release=_save))
        content.add_widget(btns)

        self.add_fam_dialog = MDDialog(
            title='Nouvelle Famille',
            type='custom',
            content_cls=content,
            size_hint=(0.90, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.add_fam_dialog)
        self.add_fam_dialog.open()
        Clock.schedule_once(lambda dt: setattr(field, 'focus', True), 0.25)

    def show_delete_family_confirmation(self, family_name: str):
        def _do_del():
            if self.db.delete_family(family_name):
                self.notify(f'Famille "{family_name}" supprimée.', 'success')
                if getattr(self, 'family_dialog', None):
                    self.family_dialog.dismiss()
                if getattr(self, 'selected_product_family_raw', '') == family_name:
                    self._select_family_and_close(AppConstants.FAM_ALL)
                self.open_family_selector_dialog()
            else:
                self.play_sound('error')
                self.notify(f'Impossible : des produits sont liés à "{family_name}".', 'error')

        self.confirm_del_fam_dialog = self._ui_confirm_dialog(
            title='Confirmer la suppression',
            message=f"Voulez-vous vraiment supprimer la famille '{self.fix_text(family_name)}' ?",
            confirm_text='SUPPRIMER',
            icon='trash-can-outline',
            confirm_icon='trash-can-outline',
            on_confirm=_do_del,
        )

    def _generate_ean13(self) -> str:
        for _ in range(20):
            base_12 = '7' + ''.join(str(random.randint(0, 9)) for _ in range(11))
            odd_sum = sum(int(base_12[i]) for i in range(1, 12, 2))
            even_sum = sum(int(base_12[i]) for i in range(0, 12, 2))
            check = (10 - (odd_sum * 3 + even_sum) % 10) % 10
            bc = f'{base_12}{check}'
            if not self.db.get_product_by_barcode(bc):
                return bc
        return '7' + ''.join(str(random.randint(0, 9)) for _ in range(12))

    # --- Fiche Produit (Ajout / Modification / Suppression) ---
    def show_manage_product_dialog(self, product, prefilled_barcode: str = ''):
        if product and AppConstants.is_default_product(product):
            self.notify('Modification interdite (produit système).', 'error')
            return

        self.temp_selected_image_path = None
        self.remove_image_order = False
        is_edit = product is not None
        has_movements = self.db.is_product_in_real_use(product['id']) if is_edit else False
        current_image = product.get('image_path', '') if is_edit else ''

        val_name = product.get('name', '') if is_edit else ''
        val_barcode = str(product.get('barcode') or '') if is_edit else (prefilled_barcode or '')
        val_ref = str(product.get('reference') or '') if is_edit else ''
        val_num = str(product.get('product_ref') or '') if is_edit else self.db.get_next_product_ref()
        raw_fam = str(product.get('family') or '') if is_edit else ''
        self.selected_product_family_raw = raw_fam or AppConstants.FAM_ALL

        s_store = FinancialEngine.to_decimal(product.get('stock_store', 0)) if is_edit else Decimal('0')
        is_unlimited = AppConstants.is_unlimited_stock(s_store) if is_edit else False
        val_stock = '' if is_unlimited else (FinancialEngine.format_qty(s_store) if is_edit else '')

        def _fmt_p(v):
            d = FinancialEngine.to_decimal(v)
            return f'{d:.2f}' if d > 0 else ''

        val_cost = _fmt_p(product.get('purchase_price', 0)) if is_edit else ''
        val_p1 = _fmt_p(product.get('price', 0)) if is_edit else ''
        val_p2 = _fmt_p(product.get('price_semi', 0)) if is_edit else ''
        val_p3 = _fmt_p(product.get('price_wholesale', 0)) if is_edit else ''

        val_is_promo = bool(int(product.get('is_promo_active', 0) or 0) == 1) if is_edit else False
        self.promo_type_selected = product.get('promo_type', 'fixed') if is_edit else 'fixed'
        val_promo_val = _fmt_p(product.get('promo_value', 0)) if is_edit else ''
        val_promo_exp = str(product.get('promo_expiry', '') or '') if is_edit else ''

        wrapper = MDBoxLayout(
            orientation='vertical',
            spacing=dp(10),
            size_hint_y=None,
            height=min(dp(610), Window.height * 0.84),
            padding=[0, dp(2), 0, 0],
        )
        scroll = MDScrollView(size_hint_y=1, bar_width=dp(3))
        main_box = MDBoxLayout(
            orientation='vertical',
            adaptive_height=True,
            spacing=dp(10),
            padding=[0, dp(2), dp(2), dp(6)],
        )

        # --- Section 1 : Informations Générales ---
        card_info = self._ui_section_card((1, 1, 1, 1), (0.88, 0.90, 0.94, 1), 'package-variant', 'Informations Générales')
        self.field_num = MDTextField(text=val_num, hint_text='N° Produit', size_hint_x=1, icon_right='pound')
        card_info.add_widget(self.field_num)

        box_bar = MDBoxLayout(orientation='horizontal', spacing=dp(8), adaptive_height=True)
        self.field_bar = MDTextField(text=val_barcode, hint_text='Code-Barres', size_hint_x=1, pos_hint={'center_y': 0.5})
        box_bar.add_widget(self.field_bar)
        box_bar.add_widget(
            self._ui_icon_pill(
                'barcode-scan',
                (0.90, 0.95, 1, 1),
                (0.09, 0.44, 0.84, 1),
                on_release=lambda x: [setattr(self, 'target_scan_field', self.field_bar), self.open_barcode_scanner(None)],
                border_color=(0.76, 0.86, 0.97, 1),
            )
        )
        box_bar.add_widget(
            self._ui_icon_pill(
                'refresh',
                (0.94, 0.95, 0.97, 1),
                (0.25, 0.30, 0.38, 1),
                on_release=lambda x: setattr(self.field_bar, 'text', self._generate_ean13()),
                border_color=(0.86, 0.88, 0.93, 1),
            )
        )
        card_info.add_widget(box_bar)

        self.field_name = SmartTextField(text=val_name, hint_text='Désignation *', required=True, icon_right='tag-text-outline')
        self.field_reference = SmartTextField(text=val_ref, hint_text='Référence (Description)', icon_right='text')
        card_info.add_widget(self.field_name)
        card_info.add_widget(self.field_reference)

        card_info.add_widget(MDLabel(text='Famille :', font_style='Caption', bold=True, theme_text_color='Secondary', size_hint_y=None, height=dp(16)))
        fam_row = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint_y=None, height=dp(46))
        fam_card = MDCard(
            size_hint_x=1,
            size_hint_y=1,
            radius=[12],
            md_bg_color=(0.95, 0.96, 0.98, 1),
            line_color=(0.86, 0.88, 0.93, 1),
            line_width=1,
            elevation=0,
            ripple_behavior=True,
            on_release=lambda x: self.open_family_selector_dialog(),
        )
        fam_inner = MDBoxLayout(orientation='horizontal', padding=[dp(12), 0, dp(10), 0], spacing=dp(8))
        fam_inner.add_widget(MDIcon(icon='shape-outline', theme_text_color='Primary', font_size='19sp', size_hint_x=None, width=dp(22), pos_hint={'center_y': 0.5}))
        self.btn_select_family = MDLabel(
            text=self.fix_text(self.selected_product_family_raw),
            font_name='ArabicFont',
            theme_text_color='Custom',
            text_color=(0.15, 0.17, 0.22, 1),
            font_style='Subtitle2',
            bold=True,
            halign='left',
            valign='center',
            shorten=True,
            shorten_from='right',
        )
        fam_inner.add_widget(self.btn_select_family)
        fam_inner.add_widget(MDIcon(icon='chevron-down', theme_text_color='Secondary', size_hint_x=None, width=dp(20), pos_hint={'center_y': 0.5}))
        fam_card.add_widget(fam_inner)
        fam_row.add_widget(fam_card)
        fam_row.add_widget(
            self._ui_icon_pill(
                'plus',
                (0.90, 0.98, 0.92, 1),
                (0.04, 0.58, 0.22, 1),
                on_release=lambda x: self.show_add_family_dialog(),
                border_color=(0.68, 0.88, 0.74, 1),
                size_dp=46,
            )
        )
        card_info.add_widget(fam_row)

        img_status = os.path.basename(current_image) if (current_image and os.path.exists(current_image)) else 'Aucune'
        img_color = (0.04, 0.58, 0.22, 1) if (current_image and os.path.exists(current_image)) else (0.45, 0.48, 0.54, 1)
        box_img = MDCard(
            orientation='horizontal',
            size_hint=(1, None),
            height=dp(48),
            radius=[12],
            elevation=0,
            md_bg_color=(0.96, 0.97, 0.99, 1),
            line_color=(0.88, 0.90, 0.94, 1),
            line_width=1,
            padding=[dp(10), 0, dp(6), 0],
            spacing=dp(8),
        )
        box_img.add_widget(MDIcon(icon='image-outline', theme_text_color='Custom', text_color=img_color, font_size='20sp', size_hint_x=None, width=dp(22), pos_hint={'center_y': 0.5}))
        self.lbl_image_status = MDLabel(
            text=f'Image : {img_status}',
            font_style='Caption',
            bold=True,
            size_hint_x=1,
            theme_text_color='Custom',
            text_color=img_color,
            shorten=True,
            shorten_from='center',
            valign='middle',
            pos_hint={'center_y': 0.5},
        )
        box_img.add_widget(self.lbl_image_status)
        box_img.add_widget(
            self._ui_icon_pill(
                'trash-can-outline',
                (1, 0.93, 0.93, 1),
                (0.82, 0.14, 0.14, 1),
                on_release=lambda x: [
                    setattr(self, 'temp_selected_image_path', None),
                    setattr(self, 'remove_image_order', True),
                    setattr(self.lbl_image_status, 'text', 'Image : Aucune'),
                    setattr(self.lbl_image_status, 'text_color', (0.45, 0.48, 0.54, 1)),
                ],
                border_color=(0.94, 0.78, 0.78, 1),
                size_dp=36,
            )
        )
        btn_pick_img = self._ui_card_btn('Choisir', (0.09, 0.44, 0.84, 1), height=36, radius=10, font_size='12.5sp', on_release=self.open_image_selector)
        btn_pick_img.size_hint_x = None
        btn_pick_img.width = dp(82)
        btn_pick_img.pos_hint = {'center_y': 0.5}
        box_img.add_widget(btn_pick_img)
        card_info.add_widget(box_img)
        main_box.add_widget(card_info)

        # --- Section 2 : Stock Initial ---
        card_stock = self._ui_section_card((0.96, 0.98, 1, 1), (0.82, 0.89, 0.97, 1), 'warehouse', 'Stock Initial', icon_color=(0.09, 0.44, 0.84, 1))
        row_stock = MDBoxLayout(orientation='horizontal', spacing=dp(10), adaptive_height=True)
        unlim_card = MDCard(
            orientation='horizontal',
            size_hint_x=0.48,
            size_hint_y=None,
            height=dp(48),
            radius=[10],
            elevation=0,
            md_bg_color=(1, 1, 1, 1),
            line_color=(0.84, 0.89, 0.96, 1),
            line_width=1,
            padding=[dp(4), 0, dp(8), 0],
            spacing=dp(2),
            pos_hint={'center_y': 0.5},
        )
        self.chk_unlimited = MDCheckbox(active=is_unlimited, size_hint=(None, None), size=(dp(38), dp(38)), pos_hint={'center_y': 0.5})
        lbl_unlim = MDLabel(
            text='Illimité\n(Verrouillé)' if has_movements else 'Illimité',
            font_style='Caption' if has_movements else 'Subtitle2',
            bold=True,
            theme_text_color='Secondary' if has_movements else 'Primary',
            valign='middle',
            pos_hint={'center_y': 0.5},
        )
        self.field_stock = MDTextField(text=val_stock, hint_text='Quantité', input_filter='float', size_hint_x=0.52, pos_hint={'center_y': 0.5})
        if has_movements:
            self.chk_unlimited.disabled = True
            self.field_stock.disabled = True
            self.field_stock.helper_text = 'Verrouillé (mouvements)'
            self.field_stock.helper_text_mode = 'persistent'
        else:
            self.chk_unlimited.bind(active=lambda c, v: setattr(self.field_stock, 'disabled', bool(v)))
            self.field_stock.disabled = is_unlimited
            unlim_card.ripple_behavior = True
            unlim_card.bind(on_release=lambda *a: setattr(self.chk_unlimited, 'active', not self.chk_unlimited.active))
        unlim_card.add_widget(self.chk_unlimited)
        unlim_card.add_widget(lbl_unlim)
        row_stock.add_widget(unlim_card)
        row_stock.add_widget(self.field_stock)
        card_stock.add_widget(row_stock)
        main_box.add_widget(card_stock)

        # --- Section 3 : Prix d'Achat ---
        card_cost = self._ui_section_card((1, 0.96, 0.96, 1), (0.94, 0.80, 0.80, 1), 'cash-minus', "Prix d'Achat (Coût / PMP)", icon_color=(0.82, 0.16, 0.16, 1))
        self.field_cost = MDTextField(text=val_cost, hint_text="Prix d'achat (DA)", input_filter='float', text_color_normal=(0.82, 0.1, 0.1, 1), text_color_focus=(0.82, 0.1, 0.1, 1))
        if has_movements:
            self.field_cost.readonly = True
            self.field_cost.helper_text = 'PMP calculé automatiquement'
            self.field_cost.helper_text_mode = 'persistent'
        card_cost.add_widget(self.field_cost)
        main_box.add_widget(card_cost)

        # --- Section 4 : Prix de Vente ---
        card_selling = self._ui_section_card((0.95, 0.99, 0.96, 1), (0.74, 0.89, 0.79, 1), 'cash-plus', 'Prix de Vente', icon_color=(0.04, 0.58, 0.22, 1))
        self.field_p1 = MDTextField(text=val_p1, hint_text='Prix Détail (Client)', input_filter='float', text_color_normal=(0, 0.52, 0.15, 1), text_color_focus=(0, 0.65, 0.2, 1))
        card_selling.add_widget(self.field_p1)
        row_p23 = MDBoxLayout(orientation='horizontal', spacing=dp(10), adaptive_height=True)
        self.field_p2 = MDTextField(text=val_p2, hint_text='Prix Demi-Gros', input_filter='float', size_hint_x=0.5, text_color_normal=(0, 0.52, 0.15, 1), text_color_focus=(0, 0.65, 0.2, 1))
        self.field_p3 = MDTextField(text=val_p3, hint_text='Prix Gros', input_filter='float', size_hint_x=0.5, text_color_normal=(0, 0.52, 0.15, 1), text_color_focus=(0, 0.65, 0.2, 1))
        row_p23.add_widget(self.field_p2)
        row_p23.add_widget(self.field_p3)
        card_selling.add_widget(row_p23)
        main_box.add_widget(card_selling)

        # --- Section 5 : Promotion ---
        self.chk_promo_active = MDCheckbox(active=val_is_promo, size_hint=(None, None), size=(dp(36), dp(36)), pos_hint={'center_y': 0.5})
        card_promo = self._ui_section_card((0.96, 0.93, 1, 1), (0.84, 0.76, 0.96, 1), 'sale', 'Promotion', icon_color=(0.75, 0.1, 0.25, 1), right_widget=self.chk_promo_active)

        type_box = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint_y=None, height=dp(40))
        self.btn_promo_fixed = MDCard(size_hint_x=0.5, size_hint_y=1, radius=[10], elevation=0, line_width=1, ripple_behavior=True)
        lbl_pf = MDLabel(text='Prix Fixe', bold=True, font_size='13sp', halign='center', valign='middle', theme_text_color='Custom')
        self.btn_promo_fixed.add_widget(lbl_pf)

        self.btn_promo_percent = MDCard(size_hint_x=0.5, size_hint_y=1, radius=[10], elevation=0, line_width=1, ripple_behavior=True)
        lbl_pp = MDLabel(text='Pourcentage %', bold=True, font_size='13sp', halign='center', valign='middle', theme_text_color='Custom')
        self.btn_promo_percent.add_widget(lbl_pp)

        self.field_promo_val = MDTextField(text=val_promo_val, hint_text='Valeur', input_filter='float')

        def _upd_promo_btns():
            is_fx = self.promo_type_selected == 'fixed'
            self.btn_promo_fixed.md_bg_color = (0.78, 0.1, 0.22, 1) if is_fx else (1, 1, 1, 1)
            self.btn_promo_fixed.line_color = (0.78, 0.1, 0.22, 1) if is_fx else (0.82, 0.76, 0.92, 1)
            lbl_pf.text_color = (1, 1, 1, 1) if is_fx else (0.25, 0.22, 0.32, 1)
            self.btn_promo_percent.md_bg_color = (1, 1, 1, 1) if is_fx else (0.78, 0.1, 0.22, 1)
            self.btn_promo_percent.line_color = (0.82, 0.76, 0.92, 1) if is_fx else (0.78, 0.1, 0.22, 1)
            lbl_pp.text_color = (0.25, 0.22, 0.32, 1) if is_fx else (1, 1, 1, 1)
            self.field_promo_val.hint_text = 'Nouveau Prix (DA)' if is_fx else 'Remise (%)'

        self.btn_promo_fixed.bind(on_release=lambda x: [setattr(self, 'promo_type_selected', 'fixed'), _upd_promo_btns()])
        self.btn_promo_percent.bind(on_release=lambda x: [setattr(self, 'promo_type_selected', 'percent'), _upd_promo_btns()])
        type_box.add_widget(self.btn_promo_fixed)
        type_box.add_widget(self.btn_promo_percent)
        card_promo.add_widget(type_box)
        card_promo.add_widget(self.field_promo_val)

        class _PromoDateCard(MDCard):
            text = StringProperty('Date Fin (Optionnel)')

            def __init__(self, **kwargs):
                super().__init__(**kwargs)
                self.orientation = 'horizontal'
                self.size_hint_x = 1
                self.size_hint_y = 1
                self.radius = [10]
                self.elevation = 0
                self.md_bg_color = (1, 1, 1, 1)
                self.line_color = (0.80, 0.72, 0.92, 1)
                self.line_width = 1
                self.padding = [dp(10), 0]
                self.spacing = dp(8)
                self.ripple_behavior = True
                self.add_widget(MDIcon(icon='calendar-outline', theme_text_color='Custom', text_color=(0.48, 0.15, 0.65, 1), font_size='19sp', size_hint_x=None, width=dp(20), pos_hint={'center_y': 0.5}))
                self._lbl = MDLabel(text=self.text, font_style='Subtitle2', bold=True, halign='left', valign='middle', theme_text_color='Custom', text_color=(0.22, 0.18, 0.30, 1), shorten=True)
                self.add_widget(self._lbl)
                self.bind(text=lambda inst, val: setattr(self._lbl, 'text', val))

        box_date = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint_y=None, height=dp(42))
        self.btn_promo_date = _PromoDateCard(text=val_promo_exp if val_promo_exp else 'Date Fin (Optionnel)')

        def _open_promo_picker(x):
            dp_dlg = MDDatePicker()
            dp_dlg.bind(on_save=lambda inst, val, dr: setattr(self.btn_promo_date, 'text', str(val)))
            dp_dlg.open()

        self.btn_promo_date.bind(on_release=_open_promo_picker)
        box_date.add_widget(self.btn_promo_date)
        box_date.add_widget(
            self._ui_icon_pill(
                'calendar-remove',
                (1, 0.93, 0.93, 1),
                (0.82, 0.14, 0.14, 1),
                on_release=lambda x: setattr(self.btn_promo_date, 'text', 'Date Fin (Optionnel)'),
                border_color=(0.94, 0.78, 0.78, 1),
                size_dp=42,
            )
        )
        card_promo.add_widget(box_date)
        _upd_promo_btns()
        main_box.add_widget(card_promo)

        scroll.add_widget(main_box)
        wrapper.add_widget(scroll)

        def _save_product(x):
            try:
                final_img = current_image
                if self.temp_selected_image_path:
                    if is_edit and current_image and os.path.exists(current_image):
                        try:
                            os.remove(current_image)
                        except Exception:
                            pass
                    final_img = self.save_product_image_local(self.temp_selected_image_path)
                elif self.remove_image_order:
                    if current_image and os.path.exists(current_image):
                        try:
                            os.remove(current_image)
                        except Exception:
                            pass
                    final_img = ''

                fam_save = '' if self.selected_product_family_raw == AppConstants.FAM_ALL else self.selected_product_family_raw
                p_exp = self.btn_promo_date.text if self.btn_promo_date.text != 'Date Fin (Optionnel)' else ''

                payload = {
                    'action': 'update' if is_edit else 'add',
                    'id': product['id'] if is_edit else None,
                    'name': self.field_name.get_value().strip(),
                    'barcode': self.field_bar.text.strip(),
                    'product_ref': self.field_num.text.strip(),
                    'reference': self.field_reference.get_value().strip(),
                    'family': fam_save,
                    'is_unlimited': self.chk_unlimited.active,
                    'stock_store': FinancialEngine.to_decimal(self.field_stock.text or '0'),
                    'purchase_price': FinancialEngine.to_decimal(self.field_cost.text or '0'),
                    'price': FinancialEngine.to_decimal(self.field_p1.text or '0'),
                    'price_semi': FinancialEngine.to_decimal(self.field_p2.text or '0'),
                    'price_wholesale': FinancialEngine.to_decimal(self.field_p3.text or '0'),
                    'tva': FinancialEngine.to_decimal(product.get('tva', 0)) if is_edit else Decimal('0'),
                    'image_path': final_img,
                    'is_promo_active': self.chk_promo_active.active,
                    'promo_type': self.promo_type_selected,
                    'promo_value': FinancialEngine.to_decimal(self.field_promo_val.text or '0'),
                    'promo_expiry': p_exp,
                }
                self.db.save_product(payload, user_name=self.current_user_name)
                if self.dialog:
                    self.dialog.dismiss()
                if hasattr(self, 'search_field') and self.search_field:
                    self.search_field.set_raw_value('')
                self.load_more_products(reset=True)
                self.check_and_load_stats()
                self.notify('Produit enregistré avec succès.', 'success')
            except Exception as exc:
                self.play_sound('error')
                self.notify(str(exc), 'error')

        def _delete_product(x):
            if has_movements:
                self.play_sound('error')
                self.notify('Suppression impossible : mouvements existants.', 'error')
                return

            def _confirm():
                try:
                    if current_image and os.path.exists(current_image):
                        try:
                            os.remove(current_image)
                        except Exception:
                            pass
                    self.db.delete_product(product['id'])
                    if self.dialog:
                        self.dialog.dismiss()
                    self.load_more_products(reset=True)
                    self.check_and_load_stats()
                    self.notify('Produit supprimé.', 'success')
                except Exception as exc:
                    self.notify(str(exc), 'error')

            self.conf_diag = self._ui_confirm_dialog(
                title='Confirmation',
                message='Voulez-vous vraiment supprimer ce produit ?',
                confirm_text='OUI, SUPPRIMER',
                cancel_text='NON',
                icon='trash-can-outline',
                confirm_icon='trash-can-outline',
                on_confirm=_confirm,
            )

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint=(1, None), height=dp(48))
        btns.add_widget(
            self._ui_card_btn(
                'FERMER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                size_hint_x=0.30 if is_edit else 0.35,
                on_release=lambda x: self.dialog.dismiss(),
            )
        )
        if is_edit:
            btns.add_widget(
                self._ui_icon_pill(
                    'trash-can-outline',
                    (1, 0.92, 0.92, 1),
                    (0.82, 0.12, 0.12, 1),
                    on_release=_delete_product,
                    border_color=(0.94, 0.74, 0.74, 1),
                    size_dp=48,
                    radius=12,
                )
            )
        btns.add_widget(
            self._ui_card_btn(
                'ENREGISTRER',
                (0.0, 0.62, 0.25, 1),
                icon='check-circle-outline',
                size_hint_x=0.70 if is_edit else 0.65,
                elevation=1,
                on_release=_save_product,
            )
        )
        wrapper.add_widget(btns)

        self.dialog = MDDialog(
            title='Fiche Produit' if is_edit else 'Nouveau Produit',
            type='custom',
            content_cls=wrapper,
            buttons=[],
            size_hint=(0.96, None),
            radius=[18, 18, 18, 18],
        )
        self._fit_dialog(self.dialog)
        self.dialog.open()
        if not is_edit:
            Clock.schedule_once(lambda dt: setattr(self.field_name, 'focus', True), 0.35)

    # --------------------------------------------------------------------------
    # 7.5 CONTRÔLE DE STOCK, PANIER & RÈGLES MÉTIER DESKTOP
    # --------------------------------------------------------------------------
    def _is_negative_stock_enforced(self) -> bool:
        doc_type = AppConstants.get_doc_type_for_mode(self.current_mode)
        if doc_type == AppConstants.DOC_TR:
            return True
        if self.db.get_bool_setting('allow_negative_stock', default=True):
            return False
        return AppConstants.STOCK_MOVEMENTS.get(doc_type, 0) == -1

    def _check_stock_availability(self, product: dict, qty_to_add: Decimal, exclude_item=None) -> tuple:
        if not self._is_negative_stock_enforced():
            return True, ''
        if AppConstants.is_default_product(product):
            return True, ''

        s_store = FinancialEngine.to_decimal(product.get('stock_store', 0))
        s_wh = FinancialEngine.to_decimal(product.get('stock_warehouse', 0))
        if AppConstants.is_unlimited_stock(s_store) or AppConstants.is_unlimited_stock(s_wh):
            return True, ''

        avail = s_wh if self.selected_location == AppConstants.LOC_WAREHOUSE else s_store
        loc_lbl = AppConstants.LOC_LABELS.get(self.selected_location, 'Magasin')

        same_edit_loc = bool(
            self.editing_tx_id
            and self.selected_location == getattr(self, 'editing_original_location', None)
        )

        pid_str = str(product.get('id'))
        in_cart = Decimal('0')
        orig_qty = Decimal('0')
        for c_item in self.cart:
            if exclude_item is not None and c_item is exclude_item:
                continue
            if str(c_item.get('id')) == pid_str:
                in_cart += FinancialEngine.to_decimal(c_item.get('qty', 0))
                if same_edit_loc:
                    orig_qty += FinancialEngine.to_decimal(c_item.get('original_server_qty', 0))
        if exclude_item is not None and same_edit_loc:
            orig_qty += FinancialEngine.to_decimal(exclude_item.get('original_server_qty', 0))

        req_total = in_cart + qty_to_add
        effective_avail = avail + orig_qty
        if req_total > effective_avail:
            return False, (
                f"Stock insuffisant ({loc_lbl}) pour '{self.fix_text(product.get('name', 'Article'))}'.\n"
                f'Disponible : {FinancialEngine.format_qty(effective_avail)} | Demandé : {FinancialEngine.format_qty(req_total)}'
            )
        return True, ''

    def _show_stock_error_dialog(self, message: str):
        if getattr(self, 'stock_error_dialog', None):
            try:
                self.stock_error_dialog.dismiss()
            except Exception:
                pass

        raw_msg = str(message or '').strip()
        main_msg = raw_msg
        avail_val = None
        req_val = None

        lines = [ln.strip() for ln in raw_msg.split('\n') if ln.strip()]
        if len(lines) >= 2 and 'Disponible' in lines[-1] and '|' in lines[-1] and 'Demandé' in lines[-1]:
            try:
                left_p, right_p = lines[-1].split('|', 1)
                avail_val = left_p.split(':', 1)[1].strip()
                req_val = right_p.split(':', 1)[1].strip()
                main_msg = '\n'.join(lines[:-1])
            except Exception:
                avail_val = None
                req_val = None

        content = MDBoxLayout(
            orientation='vertical',
            size_hint_y=None,
            adaptive_height=True,
            spacing=dp(12),
            padding=[0, 0, 0, dp(2)],
        )
        badge = self._ui_icon_pill('package-variant-closed-remove', (1, 0.92, 0.92, 1), (0.82, 0.12, 0.12, 1), size_dp=56, radius=28, icon_size='28sp')
        badge.pos_hint = {'center_x': 0.5}
        content.add_widget(badge)

        content.add_widget(MDLabel(text='Stock Insuffisant', halign='center', font_style='H6', bold=True, theme_text_color='Custom', text_color=(0.78, 0.12, 0.12, 1), size_hint_y=None, height=dp(28)))

        msg_card = MDCard(orientation='vertical', size_hint=(1, None), adaptive_height=True, radius=[12], elevation=0, md_bg_color=(1, 0.95, 0.95, 1), line_color=(0.94, 0.78, 0.78, 1), line_width=1, padding=[dp(14), dp(10)])
        msg_card.add_widget(self._ui_wrap_label(main_msg, font_style='Body2', halign='center', bold=True, theme_color='Custom', text_color=(0.25, 0.14, 0.14, 1), min_h=22))
        content.add_widget(msg_card)

        if avail_val is not None and req_val is not None:
            kpi_row = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(62))

            c_avail = MDCard(orientation='vertical', size_hint_x=0.5, radius=[12], elevation=0, md_bg_color=(0.96, 0.97, 0.99, 1), line_color=(0.86, 0.88, 0.93, 1), line_width=1, padding=[dp(8), dp(6)], spacing=dp(2))
            c_avail.add_widget(MDLabel(text='DISPONIBLE', halign='center', font_style='Caption', bold=True, theme_text_color='Secondary', size_hint_y=None, height=dp(16)))
            c_avail.add_widget(MDLabel(text=avail_val, halign='center', valign='middle', font_style='H6', bold=True, theme_text_color='Custom', text_color=(0.82, 0.12, 0.12, 1)))

            c_req = MDCard(orientation='vertical', size_hint_x=0.5, radius=[12], elevation=0, md_bg_color=(1, 0.96, 0.90, 1), line_color=(0.94, 0.80, 0.62, 1), line_width=1, padding=[dp(8), dp(6)], spacing=dp(2))
            c_req.add_widget(MDLabel(text='DEMANDÉ', halign='center', font_style='Caption', bold=True, theme_text_color='Custom', text_color=(0.75, 0.38, 0, 1), size_hint_y=None, height=dp(16)))
            c_req.add_widget(MDLabel(text=req_val, halign='center', valign='middle', font_style='H6', bold=True, theme_text_color='Custom', text_color=(0.85, 0.42, 0, 1)))

            kpi_row.add_widget(c_avail)
            kpi_row.add_widget(c_req)
            content.add_widget(kpi_row)

        content.add_widget(
            self._ui_card_btn(
                'COMPRIS',
                (0.82, 0.12, 0.12, 1),
                icon='check-circle-outline',
                font_size='15sp',
                elevation=1,
                on_release=lambda x: self.stock_error_dialog.dismiss(),
            )
        )

        self.stock_error_dialog = MDDialog(
            title='',
            type='custom',
            content_cls=content,
            size_hint=(0.90, None),
            radius=[20, 20, 20, 20],
            buttons=[],
        )
        self.stock_error_dialog.open()

    def open_add_to_cart_dialog(self, product: dict, mode: str):
        if mode == 'manage_products':
            self.show_manage_product_dialog(product)
            return

        doc_type = AppConstants.get_doc_type_for_mode(mode)
        is_qty_only = doc_type in AppConstants.QTY_ONLY_DOC_TYPES
        is_purchase = doc_type in AppConstants.PURCHASE_DOC_TYPES
        is_sale = not is_purchase and not is_qty_only
        is_virtual = AppConstants.is_default_product(product)

        ok_stk, msg_stk = self._check_stock_availability(product, Decimal('1'))
        if not ok_stk and self.sm.current == 'cart':
            self.play_sound('error')
            self._show_stock_error_dialog(msg_stk)
            return

        curr_price, has_promo, price_found_in_history = self._resolve_product_unit_price(product, doc_type)

        self.active_input_target = 'qty'
        self.input_reset_mode = True

        def _upd_colors():
            act_bg, act_border = (0.93, 0.99, 0.95, 1), (0.05, 0.62, 0.28, 1)
            inact_bg, inact_border = (0.96, 0.97, 0.98, 1), (0.86, 0.88, 0.93, 1)
            for attr, key in (('qty_card', 'qty'), ('price_card', 'price'), ('add_name_card', 'name')):
                card = getattr(self, attr, None)
                if card:
                    is_act = self.active_input_target == key
                    card.md_bg_color = act_bg if is_act else inact_bg
                    card.line_color = act_border if is_act else inact_border
                    card.line_width = 1.4 if is_act else 1.0

        header_box = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True, spacing=dp(4))
        lbl_title = self._ui_wrap_label(
            self.fix_text(product.get('name')),
            font_style='Subtitle1',
            halign='center',
            bold=True,
            theme_color='Primary',
            min_h=22,
        )
        lbl_title.font_size = '16sp'
        lbl_title.max_lines = 2
        header_box.add_widget(lbl_title)

        if has_promo:
            header_box.add_widget(MDLabel(text='PROMOTION ACTIVÉE', halign='center', theme_text_color='Custom', text_color=(0.85, 0.1, 0.1, 1), font_style='Caption', bold=True, size_hint_y=None, height=dp(18)))
        elif price_found_in_history:
            header_box.add_widget(MDLabel(text='Dernier Prix Appliqué (Historique)', halign='center', theme_text_color='Custom', text_color=(0, 0.45, 0.78, 1), font_style='Caption', bold=True, size_hint_y=None, height=dp(18)))

        base_h = 430 if is_qty_only else 510
        if is_virtual:
            base_h += 70
        target_h = min(dp(base_h), Window.height * 0.86)
        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, height=target_h, padding=[0, 0, 0, dp(2)])
        content.add_widget(header_box)

        if is_virtual:
            cnt = sum(1 for it in self.cart if AppConstants.is_default_product(it))
            default_virt_name = f'{AppConstants.DEFAULT_PRODUCT_NAME} {cnt + 1}'
            self.add_name_card = MDCard(
                size_hint=(1, None),
                height=dp(58),
                radius=[12],
                elevation=0,
                padding=[dp(12), dp(4)],
                line_width=1,
            )
            self.add_name_field = SmartTextField(
                text='',
                hint_text=f'Désignation (par défaut : {default_virt_name})',
                font_size='16sp',
                halign='center',
                mode='line',
                line_color_normal=(0, 0, 0, 0),
                line_color_focus=(0, 0, 0, 0),
                pos_hint={'center_y': 0.5},
            )
            self.add_name_field.bind(on_touch_down=lambda i, t: (setattr(self, 'active_input_target', 'name'), _upd_colors(), False)[2] if i.collide_point(*t.pos) else False)
            self.add_name_card.add_widget(self.add_name_field)
            content.add_widget(self.add_name_card)
        else:
            for attr in ('add_name_field', 'add_name_card'):
                if hasattr(self, attr):
                    delattr(self, attr)

        if not is_qty_only:
            def _select_price(*args):
                if self.active_input_target != 'price':
                    self.input_reset_mode = True
                self.active_input_target = 'price'
                _upd_colors()

            self.price_card = MDCard(
                orientation='vertical',
                size_hint=(1, None),
                height=dp(64),
                radius=[12],
                elevation=0,
                padding=[dp(12), dp(6)],
                spacing=dp(2),
                ripple_behavior=True,
                on_release=_select_price,
            )
            self.price_card.add_widget(
                MDLabel(
                    text='Dernier Prix (Historique)' if price_found_in_history else 'Prix Unitaire (DA)',
                    font_style='Caption',
                    bold=True,
                    halign='center',
                    theme_text_color='Custom',
                    text_color=(0, 0.45, 0.78, 1) if price_found_in_history else (0.45, 0.48, 0.54, 1),
                    size_hint_y=None,
                    height=dp(16),
                )
            )
            self.price_field = MDLabel(
                text=FinancialEngine.format_qty(curr_price),
                font_size='24sp',
                bold=True,
                halign='center',
                valign='middle',
                theme_text_color='Custom',
                text_color=(0, 0.45, 0.78, 1) if price_found_in_history else (0.1, 0.12, 0.16, 1),
            )
            self.price_card.add_widget(self.price_field)
            content.add_widget(self.price_card)
        else:
            for attr in ('price_field', 'price_card'):
                if hasattr(self, attr):
                    delattr(self, attr)

        def _select_qty(*args):
            if self.active_input_target != 'qty':
                self.input_reset_mode = True
            self.active_input_target = 'qty'
            _upd_colors()

        qty_row = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint_y=None, height=dp(58))
        btn_minus = self._ui_icon_pill('minus', (1, 0.93, 0.93, 1), (0.82, 0.12, 0.12, 1), size_dp=56, radius=12, icon_size='24sp', border_color=(0.94, 0.74, 0.74, 1), clickable=True)
        btn_minus.height = dp(58)

        self.qty_card = MDCard(
            orientation='vertical',
            size_hint_x=1,
            size_hint_y=None,
            height=dp(58),
            radius=[12],
            elevation=0,
            padding=[dp(10), dp(4)],
            spacing=dp(1),
            ripple_behavior=True,
            on_release=_select_qty,
        )
        self.qty_card.add_widget(MDLabel(text='Quantité', font_style='Caption', bold=True, halign='center', theme_text_color='Secondary', size_hint_y=None, height=dp(15)))
        self.qty_field = MDLabel(text='1', font_size='25sp', bold=True, halign='center', valign='middle', theme_text_color='Custom', text_color=(0.1, 0.12, 0.16, 1))
        self.qty_card.add_widget(self.qty_field)

        btn_plus = self._ui_icon_pill('plus', (0.90, 0.98, 0.92, 1), (0.04, 0.58, 0.22, 1), size_dp=56, radius=12, icon_size='24sp', border_color=(0.68, 0.88, 0.74, 1), clickable=True)
        btn_plus.height = dp(58)

        qty_row.add_widget(btn_minus)
        qty_row.add_widget(self.qty_card)
        qty_row.add_widget(btn_plus)
        content.add_widget(qty_row)

        self.btn_add = MDCard(
            orientation='vertical',
            size_hint_x=0.65,
            size_hint_y=None,
            height=dp(52),
            radius=[12],
            elevation=1,
            md_bg_color=(0.04, 0.62, 0.25, 1),
            ripple_behavior=True,
            padding=[dp(8), dp(4)],
            spacing=dp(1),
        )
        lbl_add_main = MDLabel(text='AJOUTER', bold=True, font_size='14.5sp', halign='center', valign='middle', theme_text_color='Custom', text_color=(1, 1, 1, 1))
        lbl_add_sub = MDLabel(text='', bold=True, font_size='13sp', halign='center', valign='middle', theme_text_color='Custom', text_color=(1, 1, 1, 0.92), size_hint_y=None, height=0, opacity=0)
        self.btn_add.add_widget(lbl_add_main)
        self.btn_add.add_widget(lbl_add_sub)

        def _update_btn_txt():
            if is_qty_only:
                lbl_add_sub.text = ''
                lbl_add_sub.height = 0
                lbl_add_sub.opacity = 0
                return
            q = FinancialEngine.to_decimal(self.qty_field.text or '0')
            p = FinancialEngine.to_decimal(self.price_field.text or '0')
            calc = FinancialEngine.calculate_document_totals([{'id': product['id'], 'qty': q, 'price': p, 'tva': product.get('tva', 0)}], doc_type=doc_type)
            lbl_add_sub.text = FinancialEngine.format_money(calc['final_total'])
            lbl_add_sub.height = dp(18)
            lbl_add_sub.opacity = 1

        def _confirm_add(x):
            q_val = FinancialEngine.to_decimal(self.qty_field.text or '0')
            p_val = FinancialEngine.to_decimal(self.price_field.text or '0') if not is_qty_only else Decimal('0.00')
            ok, err = FinancialEngine.validate_item_bounds(q_val, p_val, product.get('tva', 0))
            if not ok:
                self.notify(err, 'error')
                return

            custom_nm = self.add_name_field.get_value().strip() if hasattr(self, 'add_name_field') else ''
            cost_val = FinancialEngine.to_decimal(product.get('purchase_price', 0))
            if is_sale and not is_virtual and self.db.get_bool_setting('warn_price_below_cost', default=True):
                if cost_val > 0 and p_val < cost_val:
                    def _proceed_low():
                        self._finalize_add_to_cart(product, q_val, p_val, doc_type, custom_name=custom_nm)

                    warn_txt = (
                        f"Le prix de vente ({FinancialEngine.format_money(p_val)}) est inférieur au prix minimum autorisé.\n\nContinuer quand même ?"
                        if self.is_seller_mode
                        else f"Le prix de vente ({FinancialEngine.format_money(p_val)}) est inférieur au coût d'achat ({FinancialEngine.format_money(cost_val)}).\n\nContinuer quand même ?"
                    )
                    self.warn_cost_dlg = self._ui_confirm_dialog(
                        title='Attention : Prix bas',
                        message=warn_txt,
                        confirm_text='OUI, CONTINUER',
                        icon='alert-outline',
                        accent_color=(0.88, 0.45, 0.0, 1),
                        soft_bg=(1, 0.96, 0.90, 1),
                        border_color=(0.94, 0.80, 0.62, 1),
                        on_confirm=_proceed_low,
                    )
                    return
            self._finalize_add_to_cart(product, q_val, p_val, doc_type, custom_name=custom_nm)

        self.btn_add.bind(on_release=_confirm_add)

        def _delta_q(d: int):
            v = max(Decimal('1'), FinancialEngine.to_decimal(self.qty_field.text or '0') + Decimal(d))
            self.qty_field.text = FinancialEngine.format_qty(v)
            _update_btn_txt()

        btn_plus.bind(on_release=lambda x: _delta_q(1))
        btn_minus.bind(on_release=lambda x: _delta_q(-1))

        def _active_f():
            if self.active_input_target == 'name':
                return None
            if is_qty_only:
                return self.qty_field
            return self.price_field if self.active_input_target == 'price' else self.qty_field

        def _digit(d: str):
            fld = _active_f()
            if not fld:
                return
            cur = fld.text
            if self.input_reset_mode:
                fld.text = '0.' if d == '.' else d
                self.input_reset_mode = False
            elif d == '.':
                if '.' not in cur:
                    fld.text = '0.' if not cur else cur + '.'
            elif cur == '0':
                fld.text = d
            else:
                fld.text = cur + d
            _update_btn_txt()

        def _bkspc(x=None):
            fld = _active_f()
            if not fld:
                return
            self.input_reset_mode = False
            if fld.text:
                fld.text = fld.text[:-1]
            _update_btn_txt()

        grid = MDGridLayout(cols=3, spacing=dp(8), size_hint_y=1)
        for key in ('7', '8', '9', '4', '5', '6', '1', '2', '3', '.', '0', 'DEL'):
            if key == 'DEL':
                del_card = MDCard(md_bg_color=(1, 0.93, 0.93, 1), line_color=(0.94, 0.78, 0.78, 1), line_width=1, radius=[12], elevation=0, ripple_behavior=True, on_release=_bkspc)
                d_fl = MDFloatLayout()
                d_fl.add_widget(MDIcon(icon='backspace-outline', theme_text_color='Custom', text_color=(0.82, 0.14, 0.14, 1), font_size='22sp', size_hint=(None, None), size=(dp(24), dp(24)), halign='center', valign='middle', pos_hint={'center_x': 0.5, 'center_y': 0.5}))
                del_card.add_widget(d_fl)
                grid.add_widget(del_card)
            else:
                key_card = MDCard(md_bg_color=(0.96, 0.97, 0.99, 1), line_color=(0.88, 0.90, 0.94, 1), line_width=1, radius=[12], elevation=0, ripple_behavior=True, on_release=lambda x, k=key: _digit(k))
                key_card.add_widget(MDLabel(text=key, halign='center', valign='middle', font_size='21sp', bold=True, theme_text_color='Custom', text_color=(0.12, 0.14, 0.18, 1)))
                grid.add_widget(key_card)
        content.add_widget(grid)

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint_y=None, height=dp(52))
        btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.35, height=52, on_release=lambda x: self.dialog.dismiss()))
        btns.add_widget(self.btn_add)
        content.add_widget(btns)

        _upd_colors()
        _update_btn_txt()
        self.dialog = MDDialog(
            title='',
            type='custom',
            content_cls=content,
            buttons=[],
            size_hint=(0.90, None),
            radius=[20, 20, 20, 20],
        )
        self.dialog.open()

    def _finalize_add_to_cart(self, product: dict, qty: Decimal, price: Decimal, doc_type: str, custom_name: str = ''):
        ok_stk, msg_stk = self._check_stock_availability(product, qty)
        if not ok_stk:
            self.play_sound('error')
            self._show_stock_error_dialog(msg_stk)
            return

        tva_val = FinancialEngine.quantize_money(product.get('tva', 0)) if doc_type in AppConstants.INVOICE_TAX_DOC_TYPES else Decimal('0.00')
        if AppConstants.is_default_product(product):
            cnt = sum(1 for it in self.cart if AppConstants.is_default_product(it))
            new_name = custom_name.strip() if custom_name and custom_name.strip() else f'{AppConstants.DEFAULT_PRODUCT_NAME} {cnt + 1}'
            self.cart.append({
                'id': AppConstants.DEFAULT_PRODUCT_ID,
                'name': new_name,
                'price': float(FinancialEngine.quantize_money(price)),
                'qty': float(FinancialEngine.quantize_qty(qty)),
                'tva': float(tva_val),
                'discount_amount': 0.0,
                'discount_type': 'fixed',
            })
        else:
            existing = next((it for it in self.cart if int(it['id']) == int(product['id'])), None)
            if existing:
                existing['qty'] = float(FinancialEngine.quantize_qty(FinancialEngine.to_decimal(existing['qty']) + qty))
                existing['price'] = float(FinancialEngine.quantize_money(price))
            else:
                self.cart.append({
                    'id': product['id'],
                    'name': product['name'],
                    'price': float(FinancialEngine.quantize_money(price)),
                    'qty': float(FinancialEngine.quantize_qty(qty)),
                    'tva': float(tva_val),
                    'discount_amount': 0.0,
                    'discount_type': 'fixed',
                })

        if self.dialog:
            self.dialog.dismiss()
        self.update_cart_button()
        self.notify('Article ajouté au panier.', 'success')
        if hasattr(self, 'search_field') and self.search_field:
            self.search_field.set_raw_value('')
            self.load_more_products(reset=True)

    def update_cart_button(self):
        doc_type = self.editing_doc_type or AppConstants.get_doc_type_for_mode(self.current_mode)
        calc = FinancialEngine.calculate_document_totals(
            self.cart,
            doc_type=doc_type,
            global_discount=self.editing_global_discount,
        )
        if getattr(self, 'lbl_cart_count', None):
            self.lbl_cart_count.text = f'PANIER ({len(self.cart)})'
        if getattr(self, 'lbl_cart_total', None):
            self.lbl_cart_total.text = '' if doc_type in AppConstants.QTY_ONLY_DOC_TYPES else FinancialEngine.format_money(calc['final_total'])

    def remove_from_cart(self, item: dict):
        if self.editing_doc_type == AppConstants.DOC_BI:
            self.notify("Impossible de supprimer l'article d'un Bon Initial (BI).", 'warning')
            return
        if item in self.cart:
            self.cart.remove(item)
        self.refresh_cart_screen_items()
        self.update_cart_button()

    def recalculate_cart_prices(self):
        if not self.cart or not self.selected_entity:
            return
        doc_type = AppConstants.get_doc_type_for_mode(self.current_mode)
        if doc_type in AppConstants.QTY_ONLY_DOC_TYPES or doc_type == AppConstants.DOC_BI:
            return
        if doc_type in AppConstants.PURCHASE_DOC_TYPES and doc_type != AppConstants.DOC_RF:
            return
        for it in self.cart:
            if AppConstants.is_default_product(it):
                continue
            prod = self.db.get_product_by_id(it['id'])
            if prod:
                new_p, _, _ = self._resolve_product_unit_price(prod, doc_type)
                it['price'] = float(new_p)
        self.update_cart_button()
        if self.sm.current == 'cart':
            self.refresh_cart_screen_items()

    # --------------------------------------------------------------------------
    # 7.6 ÉCRAN PANIER & ÉDITION DES LIGNES
    # --------------------------------------------------------------------------
    def _build_cart_screen(self):
        class _PillSelectorCard(MDCard):
            text = StringProperty('')
            icon = StringProperty('')

            def __init__(self, **kwargs):
                super().__init__(**kwargs)
                self.orientation = 'horizontal'
                self.size_hint_y = None
                self.height = dp(44)
                self.radius = [12]
                self.elevation = 0
                self.ripple_behavior = True
                self.padding = [dp(10), 0, dp(10), 0]
                self.spacing = dp(6)

                self._ico = MDIcon(
                    icon=self.icon,
                    theme_text_color='Custom',
                    text_color=(1, 1, 1, 1),
                    font_size='19sp',
                    size_hint_x=None,
                    width=dp(22),
                    pos_hint={'center_y': 0.5},
                )
                self._lbl = MDLabel(
                    text=self.text,
                    bold=True,
                    font_size='13.5sp',
                    font_name='ArabicFont',
                    theme_text_color='Custom',
                    text_color=(1, 1, 1, 1),
                    halign='center',
                    valign='middle',
                    shorten=True,
                    shorten_from='right',
                    pos_hint={'center_y': 0.5},
                )
                self.add_widget(self._ico)
                self.add_widget(self._lbl)
                self.bind(
                    text=lambda inst, val: setattr(self._lbl, 'text', val),
                    icon=lambda inst, val: setattr(self._ico, 'icon', val),
                )

        class _ValidateCartButton(MDCard):
            text = StringProperty('')
            icon = StringProperty('check-circle-outline')

            def __init__(self, **kwargs):
                super().__init__(**kwargs)
                self.orientation = 'horizontal'
                self.size_hint = (1, None)
                self.height = dp(52)
                self.radius = [14]
                self.elevation = 2
                self.ripple_behavior = True
                self.spacing = dp(8)

                self._ico = MDIcon(
                    icon=self.icon,
                    theme_text_color='Custom',
                    text_color=(1, 1, 1, 1),
                    font_size='22sp',
                    size_hint_x=None,
                    width=dp(24),
                    pos_hint={'center_y': 0.5},
                )
                self._lbl = MDLabel(
                    text=self.text,
                    bold=True,
                    font_size='15.5sp',
                    theme_text_color='Custom',
                    text_color=(1, 1, 1, 1),
                    adaptive_width=True,
                    valign='middle',
                    pos_hint={'center_y': 0.5},
                )
                self.add_widget(MDBoxLayout())
                self.add_widget(self._ico)
                self.add_widget(self._lbl)
                self.add_widget(MDBoxLayout())
                self.bind(
                    text=lambda inst, val: setattr(self._lbl, 'text', val),
                    icon=lambda inst, val: setattr(self._ico, 'icon', val),
                )

        screen = MDScreen(name='cart')
        screen.md_bg_color = (0.95, 0.96, 0.98, 1)
        layout = MDBoxLayout(orientation='vertical')

        cart_header = MDBoxLayout(orientation='vertical', adaptive_height=True, md_bg_color=(1, 1, 1, 1))
        self.cart_toolbar = self._create_modern_header(
            'Panier',
            on_back=lambda x: self.back_to_products(),
            mode_icon='cart-check',
        )
        cart_header.add_widget(self.cart_toolbar)

        self.cart_selectors_bar = MDCard(
            orientation='horizontal',
            size_hint_y=None,
            height=dp(64),
            padding=[dp(12), dp(10)],
            spacing=dp(10),
            radius=[0],
            elevation=0,
            md_bg_color=(1, 1, 1, 1),
        )
        self.btn_ent_screen = _PillSelectorCard(
            text='Client',
            icon='account-outline',
            size_hint_x=0.5,
            md_bg_color=self.theme_cls.primary_color,
            on_release=self.handle_entity_button_click,
        )
        self.btn_loc_screen = _PillSelectorCard(
            text='Magasin',
            icon='store',
            size_hint_x=0.5,
            md_bg_color=self.theme_cls.primary_color,
            on_release=self.toggle_location,
        )
        self.cart_selectors_bar.add_widget(self.btn_ent_screen)
        self.cart_selectors_bar.add_widget(self.btn_loc_screen)
        cart_header.add_widget(self.cart_selectors_bar)
        cart_header.add_widget(MDSeparator(height=dp(1), color=(0.89, 0.91, 0.94, 1)))
        layout.add_widget(cart_header)

        self.rv_cart = CartRecycleView()
        layout.add_widget(self.rv_cart)

        self.footer_card = MDCard(
            orientation='vertical',
            size_hint_y=None,
            adaptive_height=True,
            padding=[dp(16), dp(14), dp(16), dp(14)],
            spacing=dp(10),
            radius=[22, 22, 0, 0],
            elevation=2,
            md_bg_color=(1, 1, 1, 1),
            line_color=(0.88, 0.90, 0.94, 1),
            line_width=1,
        )
        self.lbl_cart_breakdown = MDLabel(
            text='',
            font_style='Caption',
            bold=True,
            theme_text_color='Secondary',
            halign='center',
            size_hint_y=None,
            height=0,
            opacity=0,
        )
        total_row = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(36))
        self.lbl_total_title = MDLabel(text='TOTAL :', bold=True, font_style='Subtitle1', theme_text_color='Secondary', valign='middle')
        self.lbl_cart_screen_total = MDLabel(text='0,00 DA', halign='right', valign='middle', font_style='H5', bold=True, theme_text_color='Primary')
        total_row.add_widget(self.lbl_total_title)
        total_row.add_widget(self.lbl_cart_screen_total)

        self.btn_validate_cart = _ValidateCartButton(
            text='VALIDER LA COMMANDE',
            icon='check-circle-outline',
            md_bg_color=(0, 0.65, 0.25, 1),
            on_release=self.open_payment_dialog,
        )
        self.footer_card.add_widget(self.lbl_cart_breakdown)
        self.footer_card.add_widget(total_row)
        self.footer_card.add_widget(self.btn_validate_cart)
        layout.add_widget(self.footer_card)
        screen.add_widget(layout)
        return screen

    def _show_empty_cart_dialog(self, title: str = 'Panier vide', message: str = 'Veuillez ajouter au moins un produit.'):
        if getattr(self, 'dialog', None):
            try:
                self.dialog.dismiss()
            except Exception:
                pass
        content = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True, spacing=dp(12), padding=[0, 0, 0, dp(2)])
        badge = self._ui_icon_pill('cart-off', (1, 0.95, 0.88, 1), (0.88, 0.45, 0.0, 1), size_dp=56, radius=28, icon_size='28sp')
        badge.pos_hint = {'center_x': 0.5}
        content.add_widget(badge)
        content.add_widget(MDLabel(text=title, halign='center', font_style='H6', bold=True, theme_text_color='Primary', size_hint_y=None, height=dp(28)))

        msg_card = MDCard(orientation='vertical', size_hint=(1, None), adaptive_height=True, radius=[12], elevation=0, md_bg_color=(0.96, 0.97, 0.99, 1), line_color=(0.88, 0.90, 0.94, 1), line_width=1, padding=[dp(14), dp(12)])
        msg_card.add_widget(self._ui_wrap_label(message, font_style='Body2', halign='center', bold=True, theme_color='Secondary', min_h=22))
        content.add_widget(msg_card)
        content.add_widget(self._ui_card_btn('COMPRIS', (0.09, 0.44, 0.84, 1), icon='check-circle-outline', font_size='15sp', elevation=1, on_release=lambda x: self.dialog.dismiss()))

        self.dialog = MDDialog(title='', type='custom', content_cls=content, size_hint=(0.88, None), radius=[20, 20, 20, 20], buttons=[])
        self.dialog.open()

    def open_cart_screen(self, x=None):
        if not self.cart:
            self._show_empty_cart_dialog('Panier vide', 'Veuillez ajouter au moins un produit avant de continuer.')
            return
        if self.current_mode not in ('transfer', 'initial_stock') and self.selected_entity is None:
            self.show_entity_selection_dialog(None, next_action=lambda: self.open_cart_screen(None))
            return
        self.refresh_cart_screen_items()
        self.sm.transition.direction = 'left'
        self.sm.current = 'cart'

    def back_to_products(self):
        if self.editing_doc_type == AppConstants.DOC_BI:
            self.sm.transition.direction = 'right'
            self.go_back()
            return
        self.sm.transition.direction = 'right'
        self.sm.current = 'products'

    def handle_entity_button_click(self, instance):
        if self.editing_doc_type == AppConstants.DOC_BI:
            return
        if self.current_mode == 'transfer':
            self.toggle_location(instance)
        else:
            self.show_entity_selection_dialog(instance)

    def toggle_location(self, x=None):
        if self.editing_doc_type == AppConstants.DOC_BI:
            return
        self.selected_location = AppConstants.LOC_WAREHOUSE if self.selected_location == AppConstants.LOC_STORE else AppConstants.LOC_STORE
        self.update_location_display()
        if self.current_mode == 'transfer':
            src = AppConstants.LOC_LABELS[self.selected_location]
            dst = AppConstants.LOC_LABELS[AppConstants.LOC_WAREHOUSE if self.selected_location == AppConstants.LOC_STORE else AppConstants.LOC_STORE]
            txt = f'{src}  >>>  {dst}'
            if hasattr(self, 'btn_ent_screen'):
                self.btn_ent_screen.text = txt
            if hasattr(self, 'prod_toolbar'):
                self.prod_toolbar.subtitle = txt

    def update_location_display(self):
        if not hasattr(self, 'btn_loc_screen'):
            return
        if self.selected_location == AppConstants.LOC_STORE:
            self.btn_loc_screen.text = 'MAGASIN'
            self.btn_loc_screen.icon = 'store'
            self.btn_loc_screen.md_bg_color = self.theme_cls.primary_color
        else:
            self.btn_loc_screen.text = 'DÉPÔT'
            self.btn_loc_screen.icon = 'warehouse'
            self.btn_loc_screen.md_bg_color = (0.8, 0.4, 0, 1)

    def refresh_cart_screen_items(self):
        doc_type = self.editing_doc_type or AppConstants.get_doc_type_for_mode(self.current_mode)
        is_transfer = doc_type == AppConstants.DOC_TR
        is_bi = doc_type == AppConstants.DOC_BI
        is_qty_only = doc_type in AppConstants.QTY_ONLY_DOC_TYPES
        is_invoice = doc_type in AppConstants.INVOICE_TAX_DOC_TYPES
        hide_loc = is_transfer or is_bi or doc_type in (AppConstants.DOC_FP, AppConstants.DOC_DP)

        ent_id = self.selected_entity.get('id') if self.selected_entity else None
        is_comptoir = AppConstants.is_default_entity(ent_id, AppConstants.get_entity_type(doc_type))
        calc = FinancialEngine.calculate_document_totals(
            self.cart,
            doc_type=doc_type,
            global_discount=self.editing_global_discount,
            payment_method=self.editing_payment_method,
            is_comptoir=is_comptoir,
        )

        mode_colors = {
            'sale': (0.06, 0.56, 0.24, 1),
            'purchase': (0.88, 0.42, 0.0, 1),
            'invoice_sale': (0.09, 0.44, 0.84, 1),
            'invoice_purchase': (0.76, 0.14, 0.22, 1),
            'proforma': (0.5, 0.12, 0.65, 1),
            'order_purchase': (0.0, 0.5, 0.52, 1),
            'return_sale': (0.82, 0.12, 0.12, 1),
            'return_purchase': (0.0, 0.55, 0.55, 1),
            'stock_in': (0.15, 0.38, 0.75, 1),
            'stock_out': (0.48, 0.28, 0.12, 1),
            'transfer': (0.5, 0.1, 0.58, 1),
            'initial_stock': (0.0, 0.5, 0.5, 1),
        }
        hdr_col = mode_colors.get(self.current_mode, (0.09, 0.44, 0.84, 1))
        self.cart_toolbar.title = 'Bon Initial (BI)' if is_bi else f'Panier ({len(self.cart)})'
        self.cart_toolbar.subtitle = 'Modifier quantité ou prix initial' if is_bi else AppConstants.MODE_TITLES.get(self.current_mode, '')
        self.cart_toolbar.md_bg_color = hdr_col
        self.cart_toolbar.mode_icon = 'database-plus' if is_bi else ('swap-horizontal' if is_transfer else 'cart-check')
        self.update_location_display()

        if is_qty_only:
            tot_q = sum((FinancialEngine.to_decimal(it.get('qty', 0)) for it in self.cart), Decimal('0'))
            self.lbl_total_title.text = f'Résumé ({len(self.cart)} art.) :'
            self.lbl_cart_screen_total.text = f'Qté tot. : {FinancialEngine.format_qty(tot_q)}'
            self.lbl_cart_screen_total.opacity = 1
            if hasattr(self, 'lbl_cart_breakdown'):
                self.lbl_cart_breakdown.text = ''
                self.lbl_cart_breakdown.height = 0
                self.lbl_cart_breakdown.opacity = 0
        else:
            self.lbl_total_title.text = 'TOTAL (+Timbre) :' if calc['stamp_duty'] > 0 else 'TOTAL :'
            self.lbl_cart_screen_total.text = FinancialEngine.format_money(calc['final_total'])
            self.lbl_cart_screen_total.opacity = 1
            if hasattr(self, 'lbl_cart_breakdown'):
                parts = []
                if calc['total_tva_final'] > 0 or calc['stamp_duty'] > 0 or calc['total_all_discounts_final'] > 0:
                    parts.append(f"HT : {FinancialEngine.format_money(calc['total_ht_final'])}")
                if calc['total_tva_final'] > 0:
                    parts.append(f"TVA : {FinancialEngine.format_money(calc['total_tva_final'])}")
                if calc['total_all_discounts_final'] > 0:
                    parts.append(f"Remise : {FinancialEngine.format_money(calc['total_all_discounts_final'])}")
                if calc['stamp_duty'] > 0:
                    parts.append(f"Timbre : {FinancialEngine.format_money(calc['stamp_duty'])}")
                if parts:
                    self.lbl_cart_breakdown.text = '  •  '.join(parts)
                    self.lbl_cart_breakdown.height = dp(20)
                    self.lbl_cart_breakdown.opacity = 1
                else:
                    self.lbl_cart_breakdown.text = ''
                    self.lbl_cart_breakdown.height = 0
                    self.lbl_cart_breakdown.opacity = 0

        self.btn_ent_screen.disabled = is_bi
        if is_bi:
            self.btn_ent_screen.text = 'Stock Initial (Magasin • Produit verrouillé)'
            self.btn_ent_screen.icon = 'lock-outline'
            self.btn_ent_screen.md_bg_color = (0.0, 0.5, 0.5, 1)
            self.btn_validate_cart.text = 'ENREGISTRER LE BON INITIAL'
            self.btn_validate_cart.icon = 'check-circle-outline'
            self.btn_validate_cart.md_bg_color = (0.0, 0.55, 0.45, 1)
        elif is_transfer:
            src = AppConstants.LOC_LABELS[self.selected_location]
            dst = AppConstants.LOC_LABELS[AppConstants.LOC_WAREHOUSE if self.selected_location == AppConstants.LOC_STORE else AppConstants.LOC_STORE]
            self.btn_ent_screen.text = f'{src}  >>>  {dst}'
            self.btn_ent_screen.icon = 'swap-horizontal'
            self.btn_ent_screen.md_bg_color = (0.5, 0, 0.5, 1)
            self.btn_validate_cart.text = 'VALIDER LE TRANSFERT'
            self.btn_validate_cart.icon = 'swap-horizontal-bold'
            self.btn_validate_cart.md_bg_color = (0.5, 0, 0.5, 1)
        else:
            if self.selected_entity:
                self.btn_ent_screen.text = self.fix_text(self.selected_entity.get('name', 'Tiers'))
            vis = AppConstants.DOC_VISUALS.get(doc_type, {'color': (0, 0.6, 0.6, 1)})
            self.btn_ent_screen.icon = 'truck-outline' if AppConstants.get_entity_type(doc_type) == AppConstants.ENTITY_SUPPLIER else 'account-outline'
            self.btn_ent_screen.md_bg_color = vis['color']
            self.btn_validate_cart.text = 'VALIDER LA COMMANDE'
            self.btn_validate_cart.icon = 'check-circle-outline'
            self.btn_validate_cart.md_bg_color = (0, 0.65, 0.25, 1)

        if hasattr(self, 'cart_selectors_bar'):
            self.cart_selectors_bar.spacing = 0 if hide_loc else dp(10)

        if hide_loc:
            self.btn_ent_screen.size_hint_x = 1
            self.btn_loc_screen.opacity = 0
            self.btn_loc_screen.disabled = True
            self.btn_loc_screen.size_hint_x = 0
            self.btn_loc_screen.width = 0
        else:
            self.btn_ent_screen.size_hint_x = 0.5
            self.btn_loc_screen.opacity = 1
            self.btn_loc_screen.disabled = False
            self.btn_loc_screen.size_hint_x = 0.5

        rv_data = []
        for it in self.cart:
            q_disp = FinancialEngine.format_qty(it.get('qty', 0))
            b_icon = 'database-edit' if is_bi else ('pencil-plus-outline' if AppConstants.is_default_product(it) else 'package-variant')
            if is_qty_only:
                det_txt = f'Quantité : {q_disp}'
                sub_txt = ''
                d_col = [0.08, 0.42, 0.78, 1]
            else:
                line_calc = FinancialEngine.calculate_document_totals([it], doc_type=doc_type)
                p = FinancialEngine.to_decimal(it.get('price', 0))
                t_rate = FinancialEngine.to_decimal(it.get('tva', 0)) if is_invoice else Decimal('0')
                det_txt = f'{FinancialEngine.format_money(p)} x {q_disp}'
                if t_rate > 0:
                    det_txt += f' (+{int(t_rate)}% TVA)'
                sub_txt = FinancialEngine.format_money(line_calc['final_total'])
                d_col = [0.06, 0.48, 0.24, 1]
            rv_data.append({
                'name': it.get('name', 'Article'),
                'details': det_txt,
                'subtotal': sub_txt,
                'd_color': d_col,
                'badge_icon': b_icon,
                'can_delete': not is_bi,
                'raw_item': it,
            })

        self.rv_cart.data = rv_data
        self.rv_cart.refresh_from_data()

    def edit_cart_item(self, item: dict):
        doc_type = self.editing_doc_type or AppConstants.get_doc_type_for_mode(self.current_mode)
        is_invoice = doc_type in AppConstants.INVOICE_TAX_DOC_TYPES
        is_qty_only = doc_type in AppConstants.QTY_ONLY_DOC_TYPES
        is_bi = doc_type == AppConstants.DOC_BI
        is_virtual = AppConstants.is_default_product(item)

        base_h = 420 if is_qty_only else 505
        if is_virtual:
            base_h += 72
        dialog_h = min(dp(base_h), Window.height * 0.85)

        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, height=dialog_h, padding=[0, 0, 0, dp(2)])
        self.active_edit_target = 'qty'
        self.input_reset_mode = True

        def _upd_colors():
            act_bg, act_border = (0.93, 0.99, 0.95, 1), (0.05, 0.62, 0.28, 1)
            inact_bg, inact_border = (0.96, 0.97, 0.98, 1), (0.86, 0.88, 0.93, 1)
            for attr, key in (('edit_qty_card', 'qty'), ('edit_price_card', 'price'), ('edit_tva_card', 'tva'), ('edit_name_card', 'name')):
                card = getattr(self, attr, None)
                if card:
                    is_act = self.active_edit_target == key
                    card.md_bg_color = act_bg if is_act else inact_bg
                    card.line_color = act_border if is_act else inact_border
                    card.line_width = 1.4 if is_act else 1.0

        if is_virtual:
            self.edit_name_card = MDCard(size_hint=(1, None), height=dp(62), radius=[12], elevation=0, padding=[dp(12), dp(4)], line_width=1)
            self.edit_name_field = SmartTextField(
                text=item.get('name', AppConstants.DEFAULT_PRODUCT_NAME),
                hint_text="Nom de l'article",
                font_size='18sp',
                halign='center',
                mode='line',
                line_color_normal=(0, 0, 0, 0),
                line_color_focus=(0, 0, 0, 0),
                pos_hint={'center_y': 0.5},
            )
            self.edit_name_field.bind(on_touch_down=lambda i, t: (setattr(self, 'active_edit_target', 'name'), _upd_colors(), False)[2] if i.collide_point(*t.pos) else False)
            self.edit_name_card.add_widget(self.edit_name_field)
            content.add_widget(self.edit_name_card)
        else:
            for attr in ('edit_name_field', 'edit_name_card'):
                if hasattr(self, attr):
                    delattr(self, attr)
            lbl_title = self._ui_wrap_label(
                self.fix_text(item.get('name', 'Article')),
                font_style='Subtitle1',
                halign='center',
                bold=True,
                theme_color='Primary',
                min_h=22,
            )
            lbl_title.font_size = '16sp'
            lbl_title.max_lines = 2
            content.add_widget(lbl_title)

        if not is_qty_only:
            def _select_edit_price(*args):
                if self.active_edit_target != 'price':
                    self.input_reset_mode = True
                self.active_edit_target = 'price'
                _upd_colors()

            price_tva_row = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint_y=None, height=dp(62))
            self.edit_price_card = MDCard(
                orientation='vertical',
                size_hint_x=0.65 if is_invoice else 1,
                size_hint_y=None,
                height=dp(62),
                radius=[12],
                elevation=0,
                padding=[dp(10), dp(5)],
                spacing=dp(2),
                ripple_behavior=True,
                on_release=_select_edit_price,
            )
            self.edit_price_card.add_widget(
                MDLabel(
                    text="Coût d'achat initial (DA)" if is_bi else 'Prix Unitaire (DA)',
                    font_style='Caption',
                    bold=True,
                    halign='center',
                    theme_text_color='Secondary',
                    size_hint_y=None,
                    height=dp(15),
                )
            )
            self.edit_price_field = MDLabel(
                text=FinancialEngine.format_qty(item.get('price', 0)),
                font_size='23sp',
                bold=True,
                halign='center',
                valign='middle',
                theme_text_color='Custom',
                text_color=(0.1, 0.12, 0.16, 1),
            )
            self.edit_price_card.add_widget(self.edit_price_field)
            price_tva_row.add_widget(self.edit_price_card)

            if is_invoice:
                def _select_edit_tva(*args):
                    if self.active_edit_target != 'tva':
                        self.input_reset_mode = True
                    self.active_edit_target = 'tva'
                    _upd_colors()

                self.edit_tva_card = MDCard(
                    orientation='vertical',
                    size_hint_x=0.35,
                    size_hint_y=None,
                    height=dp(62),
                    radius=[12],
                    elevation=0,
                    padding=[dp(8), dp(5)],
                    spacing=dp(2),
                    ripple_behavior=True,
                    on_release=_select_edit_tva,
                )
                self.edit_tva_card.add_widget(MDLabel(text='TVA %', font_style='Caption', bold=True, halign='center', theme_text_color='Secondary', size_hint_y=None, height=dp(15)))
                self.edit_tva_field = MDLabel(text=FinancialEngine.format_qty(item.get('tva', 0)), font_size='22sp', bold=True, halign='center', valign='middle', theme_text_color='Custom', text_color=(0.1, 0.12, 0.16, 1))
                self.edit_tva_card.add_widget(self.edit_tva_field)
                price_tva_row.add_widget(self.edit_tva_card)
            else:
                for attr in ('edit_tva_field', 'edit_tva_card'):
                    if hasattr(self, attr):
                        delattr(self, attr)

            content.add_widget(price_tva_row)
        else:
            for attr in ('edit_price_field', 'edit_price_card', 'edit_tva_field', 'edit_tva_card'):
                if hasattr(self, attr):
                    delattr(self, attr)

        def _select_edit_qty(*args):
            if self.active_edit_target != 'qty':
                self.input_reset_mode = True
            self.active_edit_target = 'qty'
            _upd_colors()

        qty_row = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint_y=None, height=dp(58))
        btn_minus = self._ui_icon_pill('minus', (1, 0.93, 0.93, 1), (0.82, 0.12, 0.12, 1), size_dp=56, radius=12, icon_size='24sp', border_color=(0.94, 0.74, 0.74, 1), clickable=True)
        btn_minus.height = dp(58)

        self.edit_qty_card = MDCard(
            orientation='vertical',
            size_hint_x=1,
            size_hint_y=None,
            height=dp(58),
            radius=[12],
            elevation=0,
            padding=[dp(10), dp(4)],
            spacing=dp(1),
            ripple_behavior=True,
            on_release=_select_edit_qty,
        )
        self.edit_qty_card.add_widget(MDLabel(text='Quantité', font_style='Caption', bold=True, halign='center', theme_text_color='Secondary', size_hint_y=None, height=dp(15)))
        self.edit_qty_field = MDLabel(text=FinancialEngine.format_qty(item.get('qty', 1)), font_size='25sp', bold=True, halign='center', valign='middle', theme_text_color='Custom', text_color=(0.1, 0.12, 0.16, 1))
        self.edit_qty_card.add_widget(self.edit_qty_field)

        btn_plus = self._ui_icon_pill('plus', (0.90, 0.98, 0.92, 1), (0.04, 0.58, 0.22, 1), size_dp=56, radius=12, icon_size='24sp', border_color=(0.68, 0.88, 0.74, 1), clickable=True)
        btn_plus.height = dp(58)

        qty_row.add_widget(btn_minus)
        qty_row.add_widget(self.edit_qty_card)
        qty_row.add_widget(btn_plus)
        content.add_widget(qty_row)

        self.btn_save_edit = MDCard(
            orientation='vertical',
            size_hint_x=0.65,
            size_hint_y=None,
            height=dp(52),
            radius=[12],
            elevation=1,
            md_bg_color=(0.04, 0.62, 0.25, 1),
            ripple_behavior=True,
            padding=[dp(8), dp(4)],
            spacing=dp(1),
        )
        lbl_mod_main = MDLabel(text='MODIFIER', bold=True, font_size='14.5sp', halign='center', valign='middle', theme_text_color='Custom', text_color=(1, 1, 1, 1))
        lbl_mod_sub = MDLabel(text='', bold=True, font_size='13sp', halign='center', valign='middle', theme_text_color='Custom', text_color=(1, 1, 1, 0.92), size_hint_y=None, height=0, opacity=0)
        self.btn_save_edit.add_widget(lbl_mod_main)
        self.btn_save_edit.add_widget(lbl_mod_sub)

        def _upd_calc():
            if is_qty_only:
                lbl_mod_sub.text = ''
                lbl_mod_sub.height = 0
                lbl_mod_sub.opacity = 0
                return
            q = FinancialEngine.to_decimal(self.edit_qty_field.text or '0')
            p = FinancialEngine.to_decimal(self.edit_price_field.text or '0') if hasattr(self, 'edit_price_field') else Decimal('0')
            tva = FinancialEngine.to_decimal(self.edit_tva_field.text or '0') if hasattr(self, 'edit_tva_field') else Decimal('0')
            calc = FinancialEngine.calculate_document_totals([{'id': item['id'], 'qty': q, 'price': p, 'tva': tva}], doc_type=doc_type)
            lbl_mod_sub.text = FinancialEngine.format_money(calc['final_total'])
            lbl_mod_sub.height = dp(18)
            lbl_mod_sub.opacity = 1

        def _delta_q(d: int):
            v = max(Decimal('1'), FinancialEngine.to_decimal(self.edit_qty_field.text or '0') + Decimal(d))
            self.edit_qty_field.text = FinancialEngine.format_qty(v)
            _upd_calc()

        btn_plus.bind(on_release=lambda x: _delta_q(1))
        btn_minus.bind(on_release=lambda x: _delta_q(-1))

        def _active_f():
            if self.active_edit_target == 'name':
                return None
            if self.active_edit_target == 'price' and hasattr(self, 'edit_price_field'):
                return self.edit_price_field
            if self.active_edit_target == 'tva' and hasattr(self, 'edit_tva_field'):
                return self.edit_tva_field
            return self.edit_qty_field

        def _digit(d: str):
            fld = _active_f()
            if not fld:
                return
            cur = fld.text
            if self.input_reset_mode:
                fld.text = '0.' if d == '.' else d
                self.input_reset_mode = False
            elif d == '.':
                if '.' not in cur:
                    fld.text = '0.' if not cur else cur + '.'
            elif cur == '0':
                fld.text = d
            else:
                fld.text = cur + d
            _upd_calc()

        def _bkspc(x=None):
            fld = _active_f()
            if not fld:
                return
            self.input_reset_mode = False
            if fld.text:
                fld.text = fld.text[:-1]
            _upd_calc()

        grid = MDGridLayout(cols=3, spacing=dp(8), size_hint_y=1)
        for key in ('7', '8', '9', '4', '5', '6', '1', '2', '3', '.', '0', 'DEL'):
            if key == 'DEL':
                del_card = MDCard(md_bg_color=(1, 0.93, 0.93, 1), line_color=(0.94, 0.78, 0.78, 1), line_width=1, radius=[12], elevation=0, ripple_behavior=True, on_release=_bkspc)
                d_fl = MDFloatLayout()
                d_fl.add_widget(MDIcon(icon='backspace-outline', theme_text_color='Custom', text_color=(0.82, 0.14, 0.14, 1), font_size='22sp', size_hint=(None, None), size=(dp(24), dp(24)), halign='center', valign='middle', pos_hint={'center_x': 0.5, 'center_y': 0.5}))
                del_card.add_widget(d_fl)
                grid.add_widget(del_card)
            else:
                key_card = MDCard(md_bg_color=(0.96, 0.97, 0.99, 1), line_color=(0.88, 0.90, 0.94, 1), line_width=1, radius=[12], elevation=0, ripple_behavior=True, on_release=lambda x, k=key: _digit(k))
                key_card.add_widget(MDLabel(text=key, halign='center', valign='middle', font_size='21sp', bold=True, theme_text_color='Custom', text_color=(0.12, 0.14, 0.18, 1)))
                grid.add_widget(key_card)
        content.add_widget(grid)

        def _save_edit(x):
            new_q = FinancialEngine.to_decimal(self.edit_qty_field.text or '0')
            new_p = FinancialEngine.to_decimal(self.edit_price_field.text or '0') if hasattr(self, 'edit_price_field') else FinancialEngine.to_decimal(item.get('price', 0))
            new_tva = FinancialEngine.to_decimal(self.edit_tva_field.text or '0') if hasattr(self, 'edit_tva_field') else FinancialEngine.to_decimal(item.get('tva', 0))
            ok, err = FinancialEngine.validate_item_bounds(new_q, new_p, new_tva)
            if not ok:
                self.notify(err, 'error')
                return

            prod_db = self.db.get_product_by_id(item['id'])
            if prod_db and not is_bi:
                ok_stk, msg_stk = self._check_stock_availability(prod_db, new_q, exclude_item=item)
                if not ok_stk:
                    self.play_sound('error')
                    self._show_stock_error_dialog(msg_stk)
                    return

            item['qty'] = float(FinancialEngine.quantize_qty(new_q))
            item['price'] = float(FinancialEngine.quantize_money(new_p))
            item['tva'] = float(FinancialEngine.quantize_money(new_tva))
            if hasattr(self, 'edit_name_field'):
                nm = self.edit_name_field.get_value().strip()
                if nm:
                    item['name'] = nm

            self.refresh_cart_screen_items()
            self.update_cart_button()
            self.edit_dialog.dismiss()
            self.notify('Modifié avec succès.', 'success')

        self.btn_save_edit.bind(on_release=_save_edit)
        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint_y=None, height=dp(52))
        btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.35, height=52, on_release=lambda x: self.edit_dialog.dismiss()))
        btns.add_widget(self.btn_save_edit)
        content.add_widget(btns)

        _upd_colors()
        _upd_calc()
        self.edit_dialog = MDDialog(
            title='',
            type='custom',
            content_cls=content,
            buttons=[],
            size_hint=(0.90, None),
            radius=[20, 20, 20, 20],
        )
        self.edit_dialog.open()

    # --------------------------------------------------------------------------
    # 7.7 PAIEMENT, TIMBRE FISCAL, CRÉDIT / EXCÉDENT & VALIDATION
    # --------------------------------------------------------------------------
    def open_payment_dialog(self, x):
        now = time.time()
        if now - self._last_click_time < 0.8 or self.is_transaction_in_progress:
            return
        self._last_click_time = now

        if not self.cart:
            self._show_empty_cart_dialog('Attention', 'Le panier est vide !\nVeuillez ajouter au moins un produit.')
            return

        doc_type = self.editing_doc_type or AppConstants.get_doc_type_for_mode(self.current_mode)
        ent_id = self.selected_entity['id'] if self.selected_entity else None
        is_comptoir = AppConstants.is_default_entity(ent_id, AppConstants.get_entity_type(doc_type))
        fin_f = AppConstants.FINANCIAL_FACTORS.get(doc_type, 0)

        if self._is_negative_stock_enforced():
            same_edit_loc = bool(
                self.editing_tx_id
                and self.selected_location == getattr(self, 'editing_original_location', None)
            )
            req_by_pid = {}
            orig_by_pid = {}
            for it in self.cart:
                pid = int(it['id'])
                req_by_pid[pid] = req_by_pid.get(pid, Decimal('0')) + FinancialEngine.to_decimal(it['qty'])
                if same_edit_loc:
                    orig_by_pid[pid] = orig_by_pid.get(pid, Decimal('0')) + FinancialEngine.to_decimal(it.get('original_server_qty', 0))
            for pid, rq in req_by_pid.items():
                if AppConstants.is_default_product(pid):
                    continue
                prod_db = self.db.get_product_by_id(pid)
                if prod_db:
                    col = 'stock_warehouse' if self.selected_location == AppConstants.LOC_WAREHOUSE else 'stock_store'
                    av = FinancialEngine.to_decimal(prod_db.get(col, 0))
                    if not AppConstants.is_unlimited_stock(av) and (av + orig_by_pid.get(pid, Decimal('0'))) < rq:
                        self.play_sound('error')
                        self._show_stock_error_dialog(
                            f"Stock insuffisant ({AppConstants.LOC_LABELS[self.selected_location]}) pour '{self.fix_text(prod_db['name'])}'.\n"
                            f"Disponible : {FinancialEngine.format_qty(av + orig_by_pid.get(pid, Decimal('0')))} | Demandé : {FinancialEngine.format_qty(rq)}"
                        )
                        return

        if fin_f == 0 or doc_type in AppConstants.RETURN_DOC_TYPES or doc_type == AppConstants.DOC_BI or is_comptoir:
            method_val = (
                'Initial'
                if doc_type == AppConstants.DOC_BI
                else (
                    AppConstants.PAY_DEFAULT
                    if (fin_f == 0 or doc_type in AppConstants.RETURN_DOC_TYPES or doc_type in AppConstants.STAMP_DUTY_DOC_TYPES)
                    else AppConstants.PAY_CASH
                )
            )
            calc = FinancialEngine.calculate_document_totals(
                self.cart,
                doc_type=doc_type,
                global_discount=self.editing_global_discount,
                payment_method=method_val,
                is_comptoir=is_comptoir,
            )
            paid_val = Decimal('0.00') if (fin_f == 0 or doc_type in AppConstants.RETURN_DOC_TYPES) else calc['final_total']
            self.process_transaction(paid_val, calc['final_total'], method=method_val)
            return

        self.is_facture_tax = doc_type in AppConstants.STAMP_DUTY_DOC_TYPES
        self.show_tax_details = doc_type in AppConstants.INVOICE_TAX_DOC_TYPES
        self.payment_methods = list(AppConstants.PAYMENT_METHODS)
        self.current_method_index = 0
        if self.editing_payment_method:
            for idx, m in enumerate(self.payment_methods):
                if m['value'] == self.editing_payment_method:
                    self.current_method_index = idx
                    break

        target_h = min(dp(585 if self.show_tax_details else 550), Window.height * 0.88)
        content = MDBoxLayout(orientation='vertical', spacing=dp(9), size_hint_y=None, height=target_h, padding=[0, 0, 0, dp(2)])

        header_box = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(38), spacing=dp(8))
        header_box.add_widget(MDLabel(text='Paiement', font_style='H6', bold=True, theme_text_color='Primary', valign='center', size_hint_x=1))

        if self.is_facture_tax:
            class _MethodPill(MDCard):
                text = StringProperty('')

                def __init__(self, **kwargs):
                    super().__init__(**kwargs)
                    self.orientation = 'horizontal'
                    self.size_hint = (None, None)
                    self.size = (dp(168), dp(36))
                    self.radius = [10]
                    self.elevation = 0
                    self.md_bg_color = (0.90, 0.95, 1, 1)
                    self.line_color = (0.1, 0.45, 0.85, 0.45)
                    self.line_width = 1
                    self.ripple_behavior = True
                    self.padding = [dp(10), 0, dp(8), 0]
                    self.spacing = dp(6)
                    self.pos_hint = {'center_y': 0.5}

                    self.add_widget(MDIcon(icon='credit-card-outline', theme_text_color='Custom', text_color=(0.08, 0.4, 0.78, 1), font_size='17sp', size_hint_x=None, width=dp(18), pos_hint={'center_y': 0.5}))
                    self._lbl = MDLabel(text=self.text, bold=True, font_size='13sp', theme_text_color='Custom', text_color=(0.08, 0.35, 0.72, 1), halign='center', valign='middle', max_lines=1, shorten=False)
                    self.add_widget(self._lbl)
                    self.add_widget(MDIcon(icon='chevron-down', theme_text_color='Custom', text_color=(0.08, 0.4, 0.78, 1), font_size='17sp', size_hint_x=None, width=dp(18), pos_hint={'center_y': 0.5}))
                    self.bind(text=lambda inst, val: setattr(self._lbl, 'text', val))

            self.btn_payment_method = _MethodPill(
                text=self.payment_methods[self.current_method_index]['label'],
                on_release=self._cycle_payment_method,
            )
            header_box.add_widget(self.btn_payment_method)
        content.add_widget(header_box)

        total_card = MDCard(
            orientation='vertical',
            size_hint=(1, None),
            height=dp(106 if self.show_tax_details else 80),
            radius=[14],
            md_bg_color=(0.94, 0.97, 1, 1),
            line_color=(0.76, 0.85, 0.96, 1),
            line_width=1.1,
            elevation=0,
            padding=[dp(12), dp(8)],
            spacing=dp(4),
            ripple_behavior=True,
        )
        total_card.bind(on_release=lambda inst: [setattr(self.txt_paid, 'text', f'{self.current_final_total:.2f}'), self._recalc_ui_totals()])

        top_tot = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(18))
        top_tot.add_widget(MDLabel(text='NET À PAYER', halign='left', font_style='Caption', bold=True, theme_text_color='Custom', text_color=(0.1, 0.4, 0.75, 1)))
        top_tot.add_widget(MDLabel(text='Toucher = tout payer', halign='right', font_style='Caption', theme_text_color='Custom', text_color=(0.35, 0.52, 0.72, 1)))
        total_card.add_widget(top_tot)

        self.lbl_final_total = MDLabel(text='', halign='center', valign='middle', font_style='H5', font_size='26sp', bold=True, theme_text_color='Custom', text_color=(0.08, 0.12, 0.18, 1))
        total_card.add_widget(self.lbl_final_total)

        if self.show_tax_details:
            det_box = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(20), spacing=dp(4))
            self.lbl_pay_ht = MDLabel(text='', theme_text_color='Secondary', font_style='Caption', font_size='11sp', halign='left', bold=True, size_hint_x=0.34, shorten=True)
            self.lbl_timbre = MDLabel(text='', theme_text_color='Custom', text_color=(0.5, 0, 0.5, 1), font_style='Caption', font_size='11sp', halign='center', bold=True, size_hint_x=0.33, shorten=True)
            self.lbl_pay_tva = MDLabel(text='', theme_text_color='Custom', text_color=(0.8, 0.1, 0.1, 1), font_style='Caption', font_size='11sp', halign='right', bold=True, size_hint_x=0.33, shorten=True)
            det_box.add_widget(self.lbl_pay_ht)
            det_box.add_widget(self.lbl_timbre)
            det_box.add_widget(self.lbl_pay_tva)
            total_card.add_widget(det_box)
        else:
            self.lbl_pay_ht = None
            self.lbl_timbre = None
            self.lbl_pay_tva = None

        content.add_widget(total_card)

        paid_card = MDCard(
            orientation='vertical',
            size_hint=(1, None),
            height=dp(72),
            radius=[14],
            elevation=0,
            md_bg_color=(0.93, 0.99, 0.95, 1),
            line_color=(0.05, 0.62, 0.28, 1),
            line_width=1.4,
            padding=[dp(12), dp(6)],
            spacing=dp(2),
        )
        paid_hdr = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(18))
        paid_hdr.add_widget(MDLabel(text='MONTANT VERSÉ (DA)', font_style='Caption', bold=True, halign='center', theme_text_color='Custom', text_color=(0.04, 0.52, 0.22, 1)))
        paid_card.add_widget(paid_hdr)

        paid_val_row = MDFloatLayout(size_hint_y=1)
        self.txt_paid = MDLabel(
            text=f'{FinancialEngine.quantize_money(self.editing_payment_amount):.2f}',
            font_size='30sp',
            bold=True,
            halign='center',
            valign='middle',
            theme_text_color='Custom',
            text_color=(0.08, 0.1, 0.14, 1),
            pos_hint={'center_x': 0.5, 'center_y': 0.5},
        )
        btn_clr_paid = MDIconButton(
            icon='close-circle-outline',
            theme_text_color='Custom',
            text_color=(0.45, 0.52, 0.48, 1),
            icon_size='20sp',
            pos_hint={'right': 1.0, 'center_y': 0.5},
            on_release=lambda *a: [setattr(self.txt_paid, 'text', '0.00'), self._recalc_ui_totals()],
        )
        paid_val_row.add_widget(self.txt_paid)
        paid_val_row.add_widget(btn_clr_paid)
        paid_card.add_widget(paid_val_row)
        content.add_widget(paid_card)

        self.rest_card = MDCard(
            size_hint=(1, None),
            height=dp(36),
            radius=[10],
            elevation=0,
            md_bg_color=(1, 0.93, 0.93, 1),
            line_color=(0.94, 0.74, 0.74, 1),
            line_width=1,
            padding=[dp(10), 0],
        )
        self.lbl_rest = MDLabel(text='', halign='center', valign='middle', theme_text_color='Custom', font_style='Subtitle1', font_size='15.5sp', bold=True)
        self.rest_card.add_widget(self.lbl_rest)
        content.add_widget(self.rest_card)

        def _digit(d: str):
            cur = self.txt_paid.text
            if d == '.':
                if '.' not in cur:
                    self.txt_paid.text = '0.' if not cur else cur + '.'
            elif cur in ('0', '0.00'):
                self.txt_paid.text = d
            else:
                self.txt_paid.text = cur + d
            self._recalc_ui_totals()

        def _bkspc(inst=None):
            if self.txt_paid.text and self.txt_paid.text != '0.00':
                self.txt_paid.text = self.txt_paid.text[:-1]
            if not self.txt_paid.text:
                self.txt_paid.text = '0.00'
            self._recalc_ui_totals()

        grid = MDGridLayout(cols=3, spacing=dp(8), size_hint_y=1)
        for key in ('7', '8', '9', '4', '5', '6', '1', '2', '3', '.', '0', 'DEL'):
            if key == 'DEL':
                del_card = MDCard(md_bg_color=(1, 0.93, 0.93, 1), line_color=(0.94, 0.78, 0.78, 1), line_width=1, radius=[12], elevation=0, ripple_behavior=True, on_release=_bkspc)
                d_fl = MDFloatLayout()
                d_fl.add_widget(MDIcon(icon='backspace-outline', theme_text_color='Custom', text_color=(0.82, 0.14, 0.14, 1), font_size='22sp', size_hint=(None, None), size=(dp(24), dp(24)), halign='center', valign='middle', pos_hint={'center_x': 0.5, 'center_y': 0.5}))
                del_card.add_widget(d_fl)
                grid.add_widget(del_card)
            else:
                key_card = MDCard(md_bg_color=(0.96, 0.97, 0.99, 1), line_color=(0.88, 0.90, 0.94, 1), line_width=1, radius=[12], elevation=0, ripple_behavior=True, on_release=lambda x, k=key: _digit(k))
                key_card.add_widget(MDLabel(text=key, halign='center', valign='middle', font_size='21sp', bold=True, theme_text_color='Custom', text_color=(0.12, 0.14, 0.18, 1)))
                grid.add_widget(key_card)
        content.add_widget(grid)

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint_y=None, height=dp(52))
        btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.35, height=52, on_release=lambda x: self.pay_dialog.dismiss()))
        btns.add_widget(self._ui_card_btn('VALIDER', (0.04, 0.62, 0.25, 1), icon='check-circle-outline', size_hint_x=0.65, height=52, font_size='16sp', elevation=1, on_release=lambda x: self.finalize_submission(self.current_final_total)))
        content.add_widget(btns)

        self.pay_dialog = MDDialog(title='', type='custom', content_cls=content, buttons=[], size_hint=(0.92, None), radius=[20, 20, 20, 20])
        self._recalc_ui_totals()
        self.pay_dialog.open()

    def _cycle_payment_method(self, instance):
        self.current_method_index = (self.current_method_index + 1) % len(self.payment_methods)
        self.btn_payment_method.text = self.payment_methods[self.current_method_index]['label']
        self._recalc_ui_totals()

    def _recalc_ui_totals(self):
        doc_type = self.editing_doc_type or AppConstants.get_doc_type_for_mode(self.current_mode)
        pm = self.payment_methods[self.current_method_index]['value'] if self.is_facture_tax else AppConstants.PAY_CASH
        ent_id = self.selected_entity['id'] if self.selected_entity else None
        is_comptoir = AppConstants.is_default_entity(ent_id, AppConstants.get_entity_type(doc_type))

        calc = FinancialEngine.calculate_document_totals(
            self.cart,
            doc_type=doc_type,
            global_discount=self.editing_global_discount,
            payment_method=pm,
            is_comptoir=is_comptoir,
        )
        self.current_final_total = calc['final_total']
        if self.lbl_pay_ht:
            self.lbl_pay_ht.text = f"HT : {FinancialEngine.format_money(calc['total_ht_final'])}"
        if self.lbl_pay_tva:
            self.lbl_pay_tva.text = f"TVA : {FinancialEngine.format_money(calc['total_tva_final'])}"
        if self.lbl_timbre:
            self.lbl_timbre.text = f"Timbre : {FinancialEngine.format_money(calc['stamp_duty'])}" if calc['show_stamp_duty_widget'] else ''
            self.lbl_timbre.opacity = 1 if calc['show_stamp_duty_widget'] else 0

        self.lbl_final_total.text = FinancialEngine.format_money(self.current_final_total)
        paid = FinancialEngine.to_decimal(self.txt_paid.text or '0')
        diff = FinancialEngine.quantize_money(self.current_final_total - paid)
        if diff > 0:
            self.lbl_rest.text = f'RESTE : {FinancialEngine.format_money(diff)}'
            self.lbl_rest.text_color = (0.82, 0.1, 0.1, 1)
            if hasattr(self, 'rest_card'):
                self.rest_card.md_bg_color = (1, 0.93, 0.93, 1)
                self.rest_card.line_color = (0.94, 0.74, 0.74, 1)
        elif diff == 0:
            self.lbl_rest.text = 'RESTE : 0,00 DA'
            self.lbl_rest.text_color = (0.04, 0.56, 0.22, 1)
            if hasattr(self, 'rest_card'):
                self.rest_card.md_bg_color = (0.90, 0.98, 0.92, 1)
                self.rest_card.line_color = (0.68, 0.88, 0.74, 1)
        else:
            self.lbl_rest.text = f'EXCÉDENT / VERSEMENT : {FinancialEngine.format_money(abs(diff))}'
            self.lbl_rest.text_color = (0.04, 0.56, 0.22, 1)
            if hasattr(self, 'rest_card'):
                self.rest_card.md_bg_color = (0.90, 0.98, 0.92, 1)
                self.rest_card.line_color = (0.68, 0.88, 0.74, 1)

    def finalize_submission(self, total_amount: Decimal):
        if self.is_transaction_in_progress:
            return
        pm = self.payment_methods[self.current_method_index]['value'] if self.is_facture_tax else AppConstants.PAY_CASH
        paid_dec = FinancialEngine.quantize_money(self.txt_paid.text or '0')
        ok, err = FinancialEngine.validate_transaction_bounds(total_amount, paid_dec)
        if not ok:
            self.notify(err, 'error')
            return

        self.is_transaction_in_progress = True
        if self.pay_dialog:
            self.pay_dialog.dismiss()
            self.pay_dialog = None

        if paid_dec < total_amount:
            self.is_transaction_in_progress = False
            self.show_credit_warning(paid_dec, total_amount, total_amount - paid_dec, pm)
            return
        if paid_dec > total_amount:
            self.is_transaction_in_progress = False
            self.show_overpayment_dialog(paid_dec, total_amount, paid_dec - total_amount, pm)
            return

        Clock.schedule_once(
            lambda dt: [
                setattr(self, 'is_transaction_in_progress', False),
                self.process_transaction(paid_dec, total_amount, method=pm),
            ],
            0.1,
        )

    def show_overpayment_dialog(self, paid: Decimal, total: Decimal, excess: Decimal, method: str):
        content = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True, spacing=dp(12), padding=[0, 0, 0, dp(2)])
        badge = self._ui_icon_pill('cash-plus', (0.90, 0.98, 0.92, 1), (0.04, 0.62, 0.25, 1), size_dp=56, radius=28, icon_size='30sp')
        badge.pos_hint = {'center_x': 0.5}
        content.add_widget(badge)
        content.add_widget(MDLabel(text="Création d'un versement", halign='center', font_style='H6', bold=True, theme_text_color='Primary', size_hint_y=None, height=dp(28)))

        summary_card = MDCard(orientation='vertical', size_hint=(1, None), adaptive_height=True, radius=[12], elevation=0, md_bg_color=(0.96, 0.97, 0.99, 1), line_color=(0.88, 0.90, 0.94, 1), line_width=1, padding=[dp(14), dp(10)], spacing=dp(6))

        def _sum_row(lbl_txt: str, val_txt: str, val_col=(0.1, 0.12, 0.16, 1), lbl_col=(0.42, 0.46, 0.52, 1), is_bold: bool = False):
            r = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(22))
            r.add_widget(MDLabel(text=lbl_txt, font_style='Body2', bold=is_bold, theme_text_color='Custom', text_color=lbl_col, halign='left', valign='middle', size_hint_x=0.52))
            r.add_widget(MDLabel(text=val_txt, font_style='Body2', bold=True, theme_text_color='Custom', text_color=val_col, halign='right', valign='middle', size_hint_x=0.48))
            return r

        summary_card.add_widget(_sum_row('Montant saisi :', FinancialEngine.format_money(paid)))
        summary_card.add_widget(_sum_row('Total document :', FinancialEngine.format_money(total)))
        summary_card.add_widget(MDSeparator(height=dp(1), color=(0.86, 0.88, 0.92, 1)))
        summary_card.add_widget(_sum_row('Excédent (Versement) :', f'+ {FinancialEngine.format_money(excess)}', val_col=(0.04, 0.58, 0.22, 1), lbl_col=(0.04, 0.58, 0.22, 1), is_bold=True))
        content.add_widget(summary_card)

        note_card = MDCard(orientation='vertical', size_hint=(1, None), adaptive_height=True, radius=[10], elevation=0, md_bg_color=(0.91, 0.98, 0.93, 1), line_color=(0.72, 0.88, 0.76, 1), line_width=1, padding=[dp(12), dp(8)])
        note_card.add_widget(self._ui_wrap_label("L'excédent sera enregistré comme une opération séparée (VERSEMENT / RÈGLEMENT).", font_style='Caption', halign='center', bold=True, theme_color='Custom', text_color=(0.06, 0.45, 0.2, 1), min_h=20))
        content.add_widget(note_card)

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(48), padding=[0, dp(2), 0, 0])
        btns.add_widget(self._ui_card_btn('CORRIGER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.42, on_release=lambda x: [self.overpay_dialog.dismiss(), self.open_payment_dialog(None)]))
        btns.add_widget(self._ui_card_btn('CONFIRMER', (0.04, 0.62, 0.25, 1), icon='check-circle-outline', size_hint_x=0.58, elevation=1, on_release=lambda x: [self.overpay_dialog.dismiss(), self.process_transaction(paid, total, method=method)]))
        content.add_widget(btns)

        self.overpay_dialog = MDDialog(title='', type='custom', content_cls=content, size_hint=(0.90, None), radius=[20, 20, 20, 20], buttons=[])
        self.overpay_dialog.open()

    def show_credit_warning(self, paid: Decimal, total: Decimal, remaining: Decimal, method: str):
        content = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True, spacing=dp(12), padding=[0, 0, 0, dp(2)])
        badge = self._ui_icon_pill('alert-circle-outline', (1, 0.92, 0.92, 1), (0.82, 0.12, 0.12, 1), size_dp=56, radius=28, icon_size='30sp')
        badge.pos_hint = {'center_x': 0.5}
        content.add_widget(badge)
        content.add_widget(MDLabel(text='Attention : Crédit', halign='center', font_style='H6', bold=True, theme_text_color='Custom', text_color=(0.78, 0.12, 0.12, 1), size_hint_y=None, height=dp(28)))

        summary_card = MDCard(orientation='vertical', size_hint=(1, None), adaptive_height=True, radius=[12], elevation=0, md_bg_color=(0.96, 0.97, 0.99, 1), line_color=(0.88, 0.90, 0.94, 1), line_width=1, padding=[dp(14), dp(10)], spacing=dp(6))

        def _sum_row(lbl_txt: str, val_txt: str, val_col=(0.1, 0.12, 0.16, 1), lbl_col=(0.42, 0.46, 0.52, 1), is_bold: bool = False):
            r = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(22))
            r.add_widget(MDLabel(text=lbl_txt, font_style='Body2', bold=is_bold, theme_text_color='Custom', text_color=lbl_col, halign='left', valign='middle', size_hint_x=0.50))
            r.add_widget(MDLabel(text=val_txt, font_style='Body2', bold=True, theme_text_color='Custom', text_color=val_col, halign='right', valign='middle', size_hint_x=0.50))
            return r

        summary_card.add_widget(_sum_row('Montant saisi :', FinancialEngine.format_money(paid)))
        summary_card.add_widget(_sum_row('Total document :', FinancialEngine.format_money(total)))
        summary_card.add_widget(MDSeparator(height=dp(1), color=(0.86, 0.88, 0.92, 1)))
        summary_card.add_widget(_sum_row('Reste en Crédit :', FinancialEngine.format_money(remaining), val_col=(0.82, 0.1, 0.1, 1), lbl_col=(0.82, 0.1, 0.1, 1), is_bold=True))
        content.add_widget(summary_card)

        note_card = MDCard(orientation='vertical', size_hint=(1, None), adaptive_height=True, radius=[10], elevation=0, md_bg_color=(1, 0.94, 0.94, 1), line_color=(0.94, 0.76, 0.76, 1), line_width=1, padding=[dp(12), dp(8)])
        note_card.add_widget(self._ui_wrap_label('Le montant versé est insuffisant. Le reste sera enregistré en CRÉDIT.', font_style='Caption', halign='center', bold=True, theme_color='Custom', text_color=(0.75, 0.12, 0.12, 1), min_h=20))
        content.add_widget(note_card)

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(48), padding=[0, dp(2), 0, 0])
        btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.42, on_release=lambda x: self.debt_dialog.dismiss()))
        btns.add_widget(self._ui_card_btn('CONFIRMER', (0.82, 0.12, 0.12, 1), icon='check-circle-outline', size_hint_x=0.58, elevation=1, on_release=lambda x: [self.debt_dialog.dismiss(), self.process_transaction(paid, total, method=method)]))
        content.add_widget(btns)

        self.debt_dialog = MDDialog(title='', type='custom', content_cls=content, size_hint=(0.90, None), radius=[20, 20, 20, 20], buttons=[])
        self.debt_dialog.open()

    def process_transaction(self, paid_amount: Decimal, total_amount: Decimal, method: str = AppConstants.PAY_DEFAULT):
        if self.is_transaction_in_progress:
            return
        self.is_transaction_in_progress = True
        try:
            doc_type = self.editing_doc_type or AppConstants.get_doc_type_for_mode(self.current_mode)
            ent_cat = AppConstants.get_entity_type(doc_type)
            ent_id = self.selected_entity['id'] if self.selected_entity else (AppConstants.DEFAULT_SUPPLIER_ID if ent_cat == AppConstants.ENTITY_SUPPLIER else AppConstants.DEFAULT_CLIENT_ID)
            is_comptoir = AppConstants.is_default_entity(ent_id, ent_cat)

            calc = FinancialEngine.calculate_document_totals(
                self.cart,
                doc_type=doc_type,
                global_discount=self.editing_global_discount,
                payment_method=method,
                is_comptoir=is_comptoir,
            )
            real_total = calc['final_total']
            fin_factor = AppConstants.FINANCIAL_FACTORS.get(doc_type, 0)

            inv_paid = FinancialEngine.quantize_money(paid_amount)
            excess = Decimal('0.00')
            if doc_type == AppConstants.DOC_BI:
                inv_paid = real_total
            elif fin_factor == 1 and inv_paid > real_total:
                excess = FinancialEngine.quantize_money(inv_paid - real_total)
                inv_paid = real_total
            elif fin_factor <= 0:
                inv_paid = Decimal('0.00')

            ts = str(self.current_editing_date).split('.')[0] if (self.editing_tx_id and self.current_editing_date) else datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            tx_payload = {
                'id': self.editing_tx_id,
                'doc_type': doc_type,
                'entity_id': ent_id,
                'items': self.cart,
                'global_discount': float(self.editing_global_discount),
                'paid_amount': float(inv_paid),
                'payment_method': method,
                'location': AppConstants.LOC_STORE if doc_type == AppConstants.DOC_BI else self.selected_location,
                'note': 'Stock Initial' if doc_type == AppConstants.DOC_BI else '',
                'user_name': self.current_user_name,
                'timestamp': ts,
                'is_simple_payment': False,
            }
            saved_id = self.db.save_transaction(tx_payload)

            if excess > Decimal('0.00') and not is_comptoir:
                pay_doc = AppConstants.DOC_SUPPLIER_PAY if ent_cat == AppConstants.ENTITY_SUPPLIER else AppConstants.DOC_CLIENT_PAY
                lbl = 'Règlement' if pay_doc == AppConstants.DOC_SUPPLIER_PAY else 'Versement'
                self.db.save_transaction({
                    'doc_type': pay_doc,
                    'entity_id': ent_id,
                    'amount': float(excess),
                    'doc_number': lbl,
                    'note': lbl,
                    'user_name': self.current_user_name,
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'is_simple_payment': True,
                })

            self.notify('Opération enregistrée avec succès.', 'success')
            if self.db.get_bool_setting('printer_auto', True) and self.db.get_setting('printer_mac', '') and doc_type != AppConstants.DOC_BI:
                details = self.db.get_operation_details(saved_id, is_transfer=(doc_type == AppConstants.DOC_TR))
                if details:
                    threading.Thread(target=self.print_ticket_bluetooth, args=(details,), daemon=True).start()

            self.is_transaction_in_progress = False
            self.cart = []
            self.selected_entity = None
            self.editing_tx_id = None
            self.editing_doc_type = None
            self.editing_payment_amount = Decimal('0.00')
            self.editing_payment_method = AppConstants.PAY_DEFAULT
            self.editing_global_discount = Decimal('0.00')
            self.current_editing_date = None
            self.go_back()
            self.check_and_load_stats()
        except Exception as exc:
            self.is_transaction_in_progress = False
            self.play_sound('error')
            self._show_stock_error_dialog(str(exc))

    # --------------------------------------------------------------------------
    # 7.8 CLIENTS & FOURNISSEURS (SÉLECTION, GESTION, GPS, RIB, VERSEMENTS)
    # --------------------------------------------------------------------------
    def show_entity_selection_dialog(self, x, next_action=None):
        self.pending_entity_next_action = next_action
        e_type = AppConstants.get_entity_type(self.current_mode)
        self.entities_source_type = e_type
        title = 'Choisir un Fournisseur' if e_type == AppConstants.ENTITY_SUPPLIER else 'Choisir un Client'

        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(10),
            size_hint_y=None,
            height=min(dp(580), Window.height * 0.80),
            padding=[0, dp(2), 0, 0],
        )

        search_card = MDCard(
            orientation='horizontal',
            size_hint=(1, None),
            height=dp(50),
            radius=[12],
            elevation=0,
            md_bg_color=(0.95, 0.96, 0.98, 1),
            line_color=(0.86, 0.88, 0.93, 1),
            line_width=1,
            padding=[dp(12), 0, dp(8), 0],
            spacing=dp(8),
        )
        self.entity_search = SmartTextField(
            hint_text='Rechercher par nom ou tél...',
            mode='line',
            line_color_normal=(0, 0, 0, 0),
            line_color_focus=(0, 0, 0, 0),
            icon_right='magnify',
            pos_hint={'center_y': 0.5},
        )
        self.entity_search.bind(text=self.filter_entities_paginated)
        search_card.add_widget(self.entity_search)
        content.add_widget(search_card)

        self.rv_entity = EntityRecycleView()
        self.rv_entity.size_hint_y = 1
        content.add_widget(self.rv_entity)

        content.add_widget(
            self._ui_card_btn(
                'ANNULER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                height=46,
                on_release=lambda *a: self.entity_dialog.dismiss() if getattr(self, 'entity_dialog', None) else None,
            )
        )

        self.entity_dialog = MDDialog(
            title=title,
            type='custom',
            content_cls=content,
            size_hint=(0.96, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.entity_dialog)
        self.entity_dialog.open()
        self.active_entity_rv = self.rv_entity
        self.load_more_entities(reset=True)

    def open_entity_manager(self, entity_type: str):
        self.current_entity_type_mgmt = entity_type
        self.current_entity_sort = 'name'
        title = 'Gestion Clients' if entity_type == AppConstants.ENTITY_CLIENT else 'Gestion Fournisseurs'

        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(10),
            size_hint_y=None,
            height=min(dp(595), Window.height * 0.82),
            padding=[0, dp(2), 0, 0],
        )

        search_row = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint=(1, None), height=dp(50))
        search_card = MDCard(
            orientation='horizontal',
            size_hint_x=1,
            radius=[12],
            elevation=0,
            md_bg_color=(0.95, 0.96, 0.98, 1),
            line_color=(0.86, 0.88, 0.93, 1),
            line_width=1,
            padding=[dp(12), 0, dp(8), 0],
        )
        self.entity_search = SmartTextField(
            hint_text='Rechercher par nom ou tél...',
            mode='line',
            line_color_normal=(0, 0, 0, 0),
            line_color_focus=(0, 0, 0, 0),
            icon_right='magnify',
            pos_hint={'center_y': 0.5},
        )
        self.entity_search.bind(text=self.filter_entities_paginated)
        search_card.add_widget(self.entity_search)

        self.btn_sort_entity = self._ui_icon_pill(
            'sort-alphabetical-variant',
            (0.94, 0.95, 0.97, 1),
            (0.2, 0.2, 0.25, 1),
            on_release=self.toggle_entity_sort,
            size_dp=48,
            radius=12,
            icon_size='22sp',
            border_color=(0.86, 0.88, 0.93, 1),
        )
        search_row.add_widget(search_card)
        search_row.add_widget(self.btn_sort_entity)
        content.add_widget(search_row)

        self.rv_mgmt_entity = MgmtEntityRecycleView()
        self.rv_mgmt_entity.size_hint_y = 1
        content.add_widget(self.rv_mgmt_entity)

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(48))
        btns.add_widget(
            self._ui_card_btn(
                'FERMER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                size_hint_x=0.35,
                on_release=lambda *a: self.mgmt_dialog.dismiss() if getattr(self, 'mgmt_dialog', None) else None,
            )
        )
        btns.add_widget(
            self._ui_card_btn(
                'AJOUTER NOUVEAU',
                (0.0, 0.62, 0.25, 1),
                icon='account-plus-outline',
                size_hint_x=0.65,
                font_size='14sp',
                elevation=1,
                on_release=lambda x: self.show_add_edit_entity_dialog(None),
            )
        )
        content.add_widget(btns)

        self.mgmt_dialog = MDDialog(
            title=title,
            type='custom',
            content_cls=content,
            size_hint=(0.96, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.mgmt_dialog)
        self.mgmt_dialog.open()
        self.active_entity_rv = self.rv_mgmt_entity
        self.load_more_entities(reset=True)

    def toggle_entity_sort(self, instance):
        if self.current_entity_sort == 'name':
            self.current_entity_sort = 'balance'
            instance._ico.icon = 'sort-numeric-descending'
            instance.md_bg_color = (0.8, 0.9, 1, 1)
            self.notify('Tri : Par solde décroissant', 'info')
        elif self.current_entity_sort == 'balance':
            self.current_entity_sort = 'active'
            instance._ico.icon = 'check-bold'
            instance.md_bg_color = (0.8, 1, 0.8, 1)
            self.notify("Tri : Opérations d'aujourd'hui d'abord", 'info')
        else:
            self.current_entity_sort = 'name'
            instance._ico.icon = 'sort-alphabetical-variant'
            instance.md_bg_color = (0.94, 0.95, 0.97, 1)
            self.notify('Tri : Alphabétique (A-Z)', 'info')
        self.load_more_entities(reset=True)

    def filter_entities_paginated(self, instance, text=None):
        if self._entity_search_event:
            self._entity_search_event.cancel()
        self._entity_search_event = Clock.schedule_once(lambda dt: self.load_more_entities(reset=True), 0.25)

    def load_more_entities(self, reset: bool = False):
        if self.is_loading_entities and not reset:
            return
        is_sel = bool(getattr(self, 'entity_dialog', None) and self.entity_dialog.parent)
        is_mgmt = (not is_sel) and bool(getattr(self, 'mgmt_dialog', None) and self.mgmt_dialog.parent)
        if not is_mgmt and not is_sel:
            return

        if not reset and not self.entities_has_more:
            if self.active_entity_rv:
                self.active_entity_rv.loading_lock = False
            return

        target_type = self.current_entity_type_mgmt if is_mgmt else getattr(self, 'entities_source_type', AppConstants.ENTITY_CLIENT)
        self.active_entity_rv = self.rv_mgmt_entity if is_mgmt else self.rv_entity
        search_q = self.entity_search.get_value().strip() if hasattr(self, 'entity_search') else ''

        if reset:
            self.entity_page_offset = 0
            self.entities_has_more = True
            self.is_loading_entities = False
            if self.active_entity_rv:
                self.active_entity_rv.scroll_y = 1.0
                self.active_entity_rv.data = []

        self.is_loading_entities = True
        self._entity_fetch_seq += 1
        seq = self._entity_fetch_seq
        offset_val = self.entity_page_offset
        sort_val = self.current_entity_sort if is_mgmt else 'name'
        threading.Thread(
            target=self._load_entities_worker,
            args=(target_type, search_q, offset_val, sort_val, is_mgmt, reset, seq),
            daemon=True,
        ).start()

    def _load_entities_worker(self, target_type: str, search_q: str, offset_val: int, sort_val: str, is_mgmt: bool, reset: bool, seq: int):
        try:
            active_today = self.db.get_active_entity_ids_today(target_type)
            rows = self.db.get_entities(target_type, search_query=search_q, limit=50, offset=offset_val, sort_by=sort_val, active_ids=active_today)
            if seq != self._entity_fetch_seq:
                return

            rv_data = []
            is_client = target_type == AppConstants.ENTITY_CLIENT
            for e in rows:
                is_def = AppConstants.is_default_entity(e['id'], target_type)
                if is_mgmt and is_def:
                    continue
                bal = FinancialEngine.to_decimal(e.get('balance', 0))
                bal_str = FinancialEngine.format_money(bal)
                is_active = e['id'] in active_today
                bg_c = (0.88, 1, 0.88, 1) if is_active else (1, 1, 1, 1)
                if is_mgmt:
                    col_hex = 'D50000' if bal > 0 else '00C853'
                    rv_data.append({'raw_name': e['name'], 'balance_text': f'Solde : [color={col_hex}][b]{bal_str}[/b][/color]', 'raw_data': e, 'bg_color': bg_c})
                elif is_def:
                    rv_data.append({
                        'raw_name': 'Client' if is_client else e['name'],
                        'balance_text': '[b][color=1565C0]-- Direct (Comptoir) --[/color][/b]',
                        'icon': 'store',
                        'icon_color': [0.08, 0.42, 0.78, 1],
                        'bg_color': [0.95, 0.97, 1, 1],
                        'icon_bg': [0.86, 0.92, 1, 1],
                        'raw_data': e,
                    })
                else:
                    col_hex = 'D50000' if bal > 0 else '00A844'
                    cat = AppConstants.normalize_price_category(e.get('price_category', ''))
                    cat_suffix = f'  •  [color=555555]{cat}[/color]' if (is_client and cat != AppConstants.PRICE_CAT_DETAIL) else ''
                    rv_data.append({
                        'raw_name': e['name'],
                        'balance_text': f'Solde : [color={col_hex}][b]{bal_str}[/b][/color]{cat_suffix}',
                        'icon': 'account-check' if bal <= 0 else 'account-alert',
                        'icon_color': [0, 0.65, 0.25, 1] if bal <= 0 else [0.85, 0.1, 0.1, 1],
                        'bg_color': [1, 1, 1, 1],
                        'icon_bg': [0.91, 0.98, 0.93, 1] if bal <= 0 else [1, 0.92, 0.92, 1],
                        'raw_data': e,
                    })

            Clock.schedule_once(lambda dt: self._append_entities_to_rv(rv_data, reset, len(rows), seq), 0)
        except Exception:
            Clock.schedule_once(lambda dt: self._finish_entities_loading(), 0)

    @mainthread
    def _append_entities_to_rv(self, rv_data: list, reset: bool, db_count: int, seq: int):
        if seq != self._entity_fetch_seq:
            return
        self.entities_has_more = db_count == 50
        if self.active_entity_rv:
            if reset:
                self.entity_page_offset = db_count
            else:
                self.entity_page_offset += db_count
            self._preserve_rv_scroll_on_append(self.active_entity_rv, rv_data, reset)
        self.is_loading_entities = False

    @mainthread
    def _finish_entities_loading(self):
        self.is_loading_entities = False
        if self.active_entity_rv:
            self.active_entity_rv.loading_lock = False

    def select_entity_from_rv(self, entity_data: dict):
        self.selected_entity = {
            'id': entity_data['id'],
            'name': entity_data['name'],
            'category': AppConstants.normalize_price_category(entity_data.get('price_category', AppConstants.PRICE_CAT_DETAIL)),
        }
        if hasattr(self, 'btn_ent_screen'):
            self.btn_ent_screen.text = self.fix_text(entity_data['name'])
        if hasattr(self, 'prod_toolbar'):
            ent_lbl = 'Fournisseur' if AppConstants.get_entity_type(self.current_mode) == AppConstants.ENTITY_SUPPLIER else 'Client'
            self.prod_toolbar.subtitle = f"{ent_lbl} : {self.fix_text(entity_data['name'])}"
        self.recalculate_cart_prices()
        if getattr(self, 'entity_dialog', None):
            self.entity_dialog.dismiss()
        if getattr(self, 'pending_entity_next_action', None):
            act = self.pending_entity_next_action
            self.pending_entity_next_action = None
            act()

    def start_direct_payment_from_manager(self, entity: dict):
        self.selected_entity = entity
        self.current_mode = 'client_payment' if self.current_entity_type_mgmt == AppConstants.ENTITY_CLIENT else 'supplier_payment'
        self.editing_tx_id = None
        self.current_editing_date = None
        self.temp_note = ''
        self.show_simple_payment_dialog()

    def open_entity_edit_menu(self, entity: dict):
        if getattr(self, 'options_dialog', None):
            try:
                self.options_dialog.dismiss()
            except Exception:
                pass
        gps_data = entity.get('gps_location', '')
        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, adaptive_height=True, padding=[dp(2), dp(6), dp(2), dp(2)])

        content.add_widget(
            self._ui_card_btn(
                'Localiser (Maps)',
                (0.08, 0.62, 0.28, 1) if gps_data else (0.72, 0.74, 0.78, 1),
                icon='google-maps',
                font_size='15sp',
                on_release=lambda x: [self.options_dialog.dismiss(), self.open_client_location(gps_data)],
            )
        )

        if not self.is_seller_mode:
            row = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(46))
            row.add_widget(
                self._ui_card_btn(
                    'Modifier',
                    (0.1, 0.45, 0.85, 1),
                    icon='pencil-outline',
                    size_hint_x=0.5,
                    height=46,
                    on_release=lambda x: [self.options_dialog.dismiss(), self.show_add_edit_entity_dialog(entity)],
                )
            )
            row.add_widget(
                self._ui_card_btn(
                    'Supprimer',
                    (1, 0.93, 0.93, 1),
                    text_color=(0.85, 0.12, 0.12, 1),
                    icon='trash-can-outline',
                    size_hint_x=0.5,
                    height=46,
                    border_color=(0.88, 0.22, 0.22, 1),
                    on_release=lambda x: self.confirm_delete_entity(entity),
                )
            )
            content.add_widget(row)

        content.add_widget(
            self._ui_card_btn(
                'FERMER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                height=44,
                on_release=lambda *a: self.options_dialog.dismiss() if getattr(self, 'options_dialog', None) else None,
            )
        )

        self.options_dialog = MDDialog(
            title=self.fix_text(entity['name']),
            type='custom',
            content_cls=content,
            size_hint=(0.90, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.options_dialog)
        self.options_dialog.open()

    def open_client_location(self, location_data: str):
        loc = str(location_data or '').strip()
        if not loc:
            self.notify('Aucune position GPS enregistrée.', 'error')
            return
        url = loc if any(k in loc for k in ('http', 'waze', 'geo:')) else f'https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(loc)}'
        try:
            webbrowser.open(url)
            self.notify('Ouverture de Google Maps...', 'info')
        except Exception as exc:
            self.notify(f'Erreur : {exc}', 'error')

    def show_add_edit_entity_dialog(self, entity=None):
        is_edit = entity is not None
        is_client = self.current_entity_type_mgmt == AppConstants.ENTITY_CLIENT
        dlg_title = (
            ('Modifier Client' if is_edit else 'Nouveau Client')
            if is_client
            else ('Modifier Fournisseur' if is_edit else 'Nouveau Fournisseur')
        )

        wrapper = MDBoxLayout(
            orientation='vertical',
            spacing=dp(10),
            size_hint_y=None,
            height=min(dp(600), Window.height * 0.82),
            padding=[0, dp(2), 0, 0],
        )
        scroll = MDScrollView(size_hint_y=1, bar_width=dp(3))
        main_box = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(10), padding=[0, dp(2), dp(2), dp(6)])

        card_id = self._ui_section_card((1, 1, 1, 1), (0.88, 0.90, 0.94, 1), 'account-box-outline', 'Identité')
        f_name = SmartTextField(text=entity.get('name', '') if is_edit else '', hint_text='Nom Complet / Raison Sociale *', required=True, icon_right='account')
        f_act = SmartTextField(text=entity.get('activity', '') if is_edit else '', hint_text='Activité', icon_right='briefcase')
        card_id.add_widget(f_name)
        card_id.add_widget(f_act)
        main_box.add_widget(card_id)

        card_ct = self._ui_section_card((0.96, 0.98, 1, 1), (0.82, 0.89, 0.97, 1), 'card-account-phone-outline', 'Coordonnées & GPS')
        f_phone = SmartTextField(text=entity.get('phone', '') if is_edit else '', hint_text='Téléphone', icon_right='phone')
        f_addr = SmartTextField(text=entity.get('address', '') if is_edit else '', hint_text='Adresse', icon_right='map-marker')
        f_email = SmartTextField(text=entity.get('email', '') if is_edit else '', hint_text='Email', icon_right='email')
        gps_row = MDBoxLayout(orientation='horizontal', spacing=dp(6), adaptive_height=True)
        f_gps = SmartTextField(text=entity.get('gps_location', '') if is_edit else '', hint_text='Lien Maps / Position GPS', icon_right='google-maps', size_hint_x=1)
        gps_row.add_widget(f_gps)
        gps_row.add_widget(
            self._ui_icon_pill(
                'content-paste',
                (0.90, 0.95, 1, 1),
                self.theme_cls.primary_color,
                on_release=lambda x: f_gps.set_raw_value(Clipboard.paste()),
                size_dp=38,
                icon_size='19sp',
            )
        )
        for w in (f_phone, f_addr, gps_row, f_email):
            card_ct.add_widget(w)
        main_box.add_widget(card_ct)

        disp_cat = AppConstants.normalize_price_category(entity.get('price_category', AppConstants.PRICE_CAT_DETAIL) if is_edit else AppConstants.PRICE_CAT_DETAIL)
        f_cat = MDTextField(text=disp_cat, hint_text='Catégorie de Prix', readonly=True, icon_right='menu-down')
        f_cat.bind(on_touch_down=lambda i, t: (self.show_price_cat_selector(i), True)[1] if i.collide_point(*t.pos) else False)
        if is_client:
            card_cm = self._ui_section_card((1, 0.98, 0.94, 1), (0.94, 0.87, 0.76, 1), 'store-cog-outline', 'Commercial', icon_color=(0.8, 0.45, 0, 1))
            card_cm.add_widget(f_cat)
            main_box.add_widget(card_cm)

        card_fisc = self._ui_section_card((0.97, 0.97, 0.98, 1), (0.88, 0.89, 0.92, 1), 'file-document-multiple-outline', 'Informations Fiscales & Bancaires')
        f_rc = SmartTextField(text=entity.get('rc', '') if is_edit else '', hint_text='N° Registre Commerce (RC)')
        f_nif = SmartTextField(text=entity.get('nif', '') if is_edit else '', hint_text='N.I.F')
        f_nis = SmartTextField(text=entity.get('nis', '') if is_edit else '', hint_text='N.I.S')
        f_nai = SmartTextField(text=entity.get('nai', '') if is_edit else '', hint_text='N.A.I')
        f_rib = SmartTextField(text=entity.get('rib', '') if is_edit else '', hint_text='R.I.B')
        for w in (f_rc, f_nif, f_nis, f_nai, f_rib):
            card_fisc.add_widget(w)
        main_box.add_widget(card_fisc)

        scroll.add_widget(main_box)
        wrapper.add_widget(scroll)

        def _save(x):
            try:
                self.db.save_entity({
                    'action': 'update' if is_edit else 'add',
                    'id': entity['id'] if is_edit else None,
                    'type': self.current_entity_type_mgmt,
                    'name': f_name.get_value().strip(),
                    'activity': f_act.get_value().strip(),
                    'phone': f_phone.get_value().strip(),
                    'address': f_addr.get_value().strip(),
                    'email': f_email.get_value().strip(),
                    'gps_location': f_gps.get_value().strip(),
                    'price_category': f_cat.text,
                    'rc': f_rc.get_value().strip(),
                    'nif': f_nif.get_value().strip(),
                    'nis': f_nis.get_value().strip(),
                    'nai': f_nai.get_value().strip(),
                    'rib': f_rib.get_value().strip(),
                })
                self.ae_dialog.dismiss()
                self.notify('Enregistré avec succès.', 'success')
                self.load_more_entities(reset=True)
            except Exception as exc:
                self.notify(str(exc), 'error')

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(48))
        btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.35, on_release=lambda *a: self.ae_dialog.dismiss() if getattr(self, 'ae_dialog', None) else None))
        btns.add_widget(self._ui_card_btn('ENREGISTRER', (0.0, 0.62, 0.25, 1), icon='check-circle-outline', size_hint_x=0.65, font_size='15sp', elevation=1, on_release=_save))
        wrapper.add_widget(btns)

        self.ae_dialog = MDDialog(
            title=dlg_title,
            type='custom',
            content_cls=wrapper,
            size_hint=(0.96, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.ae_dialog)
        self.ae_dialog.open()
        Clock.schedule_once(lambda dt: setattr(f_name, 'focus', True), 0.3)

    def show_price_cat_selector(self, tf):
        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, adaptive_height=True, padding=[dp(4), dp(6), dp(4), dp(4)])
        for cat, col in (
            (AppConstants.PRICE_CAT_DETAIL, (0.0, 0.55, 0.58, 1)),
            (AppConstants.PRICE_CAT_SEMI, (0.9, 0.52, 0.0, 1)),
            (AppConstants.PRICE_CAT_GROS, (0.48, 0.1, 0.58, 1)),
        ):
            content.add_widget(
                self._ui_card_btn(
                    cat,
                    col,
                    icon='tag-outline',
                    height=46,
                    font_size='15sp',
                    elevation=1,
                    on_release=lambda x, v=cat: [setattr(tf, 'text', v), self.cat_dialog.dismiss()],
                )
            )
        self.cat_dialog = MDDialog(title='Catégorie de Prix', type='custom', content_cls=content, size_hint=(0.82, None), radius=[18, 18, 18, 18])
        self._fit_dialog(self.cat_dialog)
        self.cat_dialog.open()

    def confirm_delete_entity(self, entity: dict):
        if getattr(self, 'options_dialog', None):
            self.options_dialog.dismiss()

        def _do_del():
            try:
                if self.db.delete_entity(entity['id']):
                    self.notify('Supprimé avec succès.', 'success')
                    self.load_more_entities(reset=True)
                else:
                    self.play_sound('error')
                    self.notify('Suppression impossible : ce compte possède des opérations.', 'error')
            except Exception as exc:
                self.notify(str(exc), 'error')

        self.del_conf_dialog = self._ui_confirm_dialog(
            title='Confirmer la suppression',
            message=f"Voulez-vous vraiment supprimer '{self.fix_text(entity['name'])}' ?",
            confirm_text='SUPPRIMER',
            icon='account-remove-outline',
            confirm_icon='trash-can-outline',
            on_confirm=_do_del,
        )

    def show_simple_payment_dialog(self, amount=None):
        if not self.selected_entity:
            return
        self.temp_note = getattr(self, 'temp_note', '')
        is_edit = self.editing_tx_id is not None
        is_client = self.current_mode in ('client_payment', 'client_pay')
        title = 'Versement Client' if is_client else 'Règlement Fournisseur'
        theme_col = (0.0, 0.62, 0.25, 1) if is_client else (0.85, 0.42, 0.0, 1)
        soft_bg = (0.95, 0.99, 0.96, 1) if is_client else (1.0, 0.97, 0.93, 1)

        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, adaptive_height=True, padding=[dp(2), dp(6), dp(2), dp(2)])

        header_box = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True, spacing=dp(2))
        lbl_ent = self._ui_wrap_label(self.fix_text(self.selected_entity['name']), font_style='H6', halign='center', bold=True, theme_color='Primary', min_h=26)
        lbl_ent.max_lines = 2
        header_box.add_widget(lbl_ent)
        if 'balance' in self.selected_entity:
            cur_bal = FinancialEngine.to_decimal(self.selected_entity.get('balance', 0))
            bal_col = (0.82, 0.1, 0.1, 1) if cur_bal > 0 else (0.0, 0.6, 0.2, 1)
            header_box.add_widget(MDLabel(text=f'Solde actuel : {FinancialEngine.format_money(cur_bal)}', halign='center', font_style='Caption', bold=True, theme_text_color='Custom', text_color=bal_col, size_hint_y=None, height=dp(18)))
        content.add_widget(header_box)

        val_str = FinancialEngine.format_qty(amount) if amount is not None else ''
        amount_card = MDCard(orientation='vertical', size_hint=(1, None), height=dp(92), radius=[14], elevation=0, md_bg_color=soft_bg, line_color=theme_col, line_width=1.3, padding=[dp(14), dp(8), dp(14), dp(8)])
        self.txt_simple_amount = NoMenuTextField(text=val_str, hint_text='Montant (DA)', font_size='38sp', halign='center', readonly=True, mode='line', line_color_normal=(0, 0, 0, 0), line_color_focus=(0, 0, 0, 0), pos_hint={'center_y': 0.5})
        self.txt_simple_amount.theme_text_color = 'Custom'

        def _sync_amount_color():
            is_neg = self.txt_simple_amount.text.startswith('-')
            col = (0.82, 0.1, 0.1, 1) if is_neg else (0.08, 0.1, 0.12, 1)
            self.txt_simple_amount.text_color_normal = col
            self.txt_simple_amount.text_color_focus = col

        _sync_amount_color()
        amount_card.add_widget(self.txt_simple_amount)
        content.add_widget(amount_card)

        def _bkspc(x=None):
            if self.txt_simple_amount.text:
                self.txt_simple_amount.text = self.txt_simple_amount.text[:-1]
            _sync_amount_color()

        def _clear_all(x=None):
            self.txt_simple_amount.text = ''
            _sync_amount_color()

        def _digit(d: str):
            cur = self.txt_simple_amount.text
            if d == '.':
                if '.' not in cur:
                    self.txt_simple_amount.text = '0.' if not cur else ('-0.' if cur == '-' else cur + '.')
            elif d == '-':
                self.txt_simple_amount.text = cur[1:] if cur.startswith('-') else '-' + cur
            elif cur == '0':
                self.txt_simple_amount.text = d
            elif cur == '-0':
                self.txt_simple_amount.text = '-' + d
            else:
                self.txt_simple_amount.text = cur + d
            _sync_amount_color()

        tools_row = MDBoxLayout(orientation='horizontal', size_hint=(1, None), height=dp(42), spacing=dp(8))
        self.btn_note_icon = self._ui_card_btn(
            'Note',
            (0, 0.6, 0.2, 1) if self.temp_note else (0.22, 0.24, 0.28, 1),
            icon='note-edit-outline',
            size_hint_x=0.52,
            height=42,
            radius=10,
            on_release=self.open_note_input,
        )
        btn_clear = self._ui_card_btn(
            'C',
            (0.92, 0.93, 0.95, 1),
            text_color=(0.25, 0.25, 0.3, 1),
            size_hint_x=0.24,
            height=42,
            radius=10,
            font_size='18sp',
            on_release=_clear_all,
        )
        btn_del = MDCard(
            size_hint_x=0.24,
            size_hint_y=None,
            height=dp(42),
            radius=[10],
            elevation=0,
            md_bg_color=(0.9, 0.15, 0.15, 1),
            ripple_behavior=True,
            on_release=_bkspc,
        )
        del_fl = MDFloatLayout()
        del_fl.add_widget(
            MDIcon(
                icon='backspace-outline',
                theme_text_color='Custom',
                text_color=(1, 1, 1, 1),
                font_size='22sp',
                size_hint=(None, None),
                size=(dp(24), dp(24)),
                halign='center',
                valign='middle',
                pos_hint={'center_x': 0.5, 'center_y': 0.5},
            )
        )
        btn_del.add_widget(del_fl)

        tools_row.add_widget(self.btn_note_icon)
        tools_row.add_widget(btn_clear)
        tools_row.add_widget(btn_del)
        content.add_widget(tools_row)

        grid_h = min(dp(236), max(dp(188), Window.height * 0.29))
        grid = MDGridLayout(cols=3, spacing=dp(8), size_hint=(1, None), height=grid_h)
        for key in ('7', '8', '9', '4', '5', '6', '1', '2', '3', '-', '0', '.'):
            if key == '-':
                btn = MDCard(
                    orientation='vertical',
                    md_bg_color=(1, 0.89, 0.89, 1),
                    line_color=(0.92, 0.62, 0.62, 1),
                    line_width=1,
                    elevation=1,
                    radius=[12],
                    ripple_behavior=True,
                    padding=[dp(4), dp(6), dp(4), dp(6)],
                    spacing=dp(2),
                    on_release=lambda x, k=key: _digit(k),
                )
                btn.add_widget(MDBoxLayout())
                btn.add_widget(
                    MDIcon(
                        icon='minus-thick',
                        font_size='22sp',
                        size_hint=(None, None),
                        size=(dp(24), dp(24)),
                        halign='center',
                        valign='middle',
                        pos_hint={'center_x': 0.5},
                        theme_text_color='Custom',
                        text_color=(0.8, 0, 0, 1),
                    )
                )
                btn.add_widget(MDLabel(text='Crédit Initial', font_style='Caption', bold=True, halign='center', theme_text_color='Custom', text_color=(0.8, 0, 0, 1), adaptive_height=True))
                btn.add_widget(MDBoxLayout())
            else:
                btn = MDCard(
                    md_bg_color=(0.97, 0.97, 0.98, 1),
                    line_color=(0.88, 0.89, 0.92, 1),
                    line_width=1,
                    elevation=1,
                    radius=[12],
                    ripple_behavior=True,
                    on_release=lambda x, k=key: _digit(k),
                )
                btn.add_widget(MDLabel(text=key, halign='center', valign='middle', font_size='25sp', bold=True, theme_text_color='Custom', text_color=(0.12, 0.14, 0.18, 1)))
            grid.add_widget(btn)
        content.add_widget(grid)

        def _cancel_simple_pay(*args):
            self.editing_tx_id = None
            self.current_editing_date = None
            self.temp_note = ''
            if getattr(self, 'simple_pay_dialog', None):
                self.simple_pay_dialog.dismiss()

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint=(1, None), height=dp(52), padding=[0, dp(4), 0, 0])
        btns.add_widget(
            self._ui_card_btn(
                'ANNULER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                size_hint_x=0.30 if is_edit else 0.34,
                height=52,
                on_release=_cancel_simple_pay,
            )
        )
        if is_edit:
            btns.add_widget(
                self._ui_icon_pill(
                    'trash-can-outline',
                    (1, 0.92, 0.92, 1),
                    (0.82, 0.12, 0.12, 1),
                    on_release=lambda x: self.confirm_delete_simple_payment(),
                    border_color=(0.94, 0.74, 0.74, 1),
                    size_dp=50,
                    radius=12,
                )
            )
        btns.add_widget(
            self._ui_card_btn(
                'VALIDER',
                theme_col,
                icon='check-circle-outline',
                size_hint_x=0.70 if is_edit else 0.66,
                height=52,
                font_size='16sp',
                elevation=2,
                on_release=self.submit_simple_payment,
            )
        )
        content.add_widget(btns)

        self.simple_pay_dialog = MDDialog(
            title=title,
            type='custom',
            content_cls=content,
            size_hint=(0.94, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.simple_pay_dialog)
        self.simple_pay_dialog.open()

    def confirm_delete_simple_payment(self):
        if not self.editing_tx_id:
            return
        tx_id_del = self.editing_tx_id

        def _do_del():
            if getattr(self, 'simple_pay_dialog', None):
                self.simple_pay_dialog.dismiss()
            try:
                self.db.delete_transaction(tx_id_del, is_transfer=False)
                self.editing_tx_id = None
                self.current_editing_date = None
                self.temp_note = ''
                self.notify('Opération financière supprimée.', 'success')
                self.load_more_entities(reset=True)
                self.check_and_load_stats()
                if getattr(self, 'pending_dialog', None) and self.pending_dialog.parent:
                    self.history_page_offset = 0
                    self.history_has_more = True
                    self.is_loading_history = False
                    if hasattr(self, 'rv_history'):
                        self.rv_history.data = []
                        self.rv_history.scroll_y = 1.0
                    self.load_more_history()
                if getattr(self, 'entity_hist_dialog', None) and self.entity_hist_dialog.parent:
                    if getattr(self, 'history_target_entity', None) and self.history_target_entity.get('id'):
                        fresh_ent = self.db.get_entity_by_id(self.history_target_entity['id'])
                        if fresh_ent:
                            self.history_target_entity = fresh_ent
                    self._reload_entity_history_ui()
            except Exception as exc:
                self.notify(str(exc), 'error')

        self._ui_confirm_dialog(
            title='Confirmer la suppression',
            message='Voulez-vous vraiment supprimer cette opération financière ?\nLe solde sera recalculé automatiquement.',
            confirm_text='SUPPRIMER',
            icon='trash-can-outline',
            confirm_icon='trash-can-outline',
            on_confirm=_do_del,
        )

    def open_note_input(self, instance):
        content = MDBoxLayout(orientation='vertical', spacing=dp(12), size_hint_y=None, adaptive_height=True, padding=[0, dp(4), 0, 0])
        note_card = MDCard(orientation='vertical', size_hint=(1, None), height=dp(70), radius=[12], elevation=0, md_bg_color=(0.96, 0.97, 0.99, 1), line_color=(0.86, 0.88, 0.93, 1), line_width=1, padding=[dp(12), dp(4)])
        nf = SmartTextField(text=self.temp_note, hint_text='Note (optionnel)', icon_right='note-text-outline', pos_hint={'center_y': 0.5})
        note_card.add_widget(nf)
        content.add_widget(note_card)

        def _save_n(x):
            self.temp_note = nf.get_value().strip()
            self.btn_note_icon.md_bg_color = (0, 0.6, 0.2, 1) if self.temp_note else (0.22, 0.24, 0.28, 1)
            ndlg.dismiss()

        def _clear_n(x):
            self.temp_note = ''
            self.btn_note_icon.md_bg_color = (0.22, 0.24, 0.28, 1)
            ndlg.dismiss()

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint=(1, None), height=dp(46))
        btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.35, height=46, on_release=lambda x: ndlg.dismiss()))
        btns.add_widget(self._ui_icon_pill('trash-can-outline', (1, 0.93, 0.93, 1), (0.82, 0.14, 0.14, 1), on_release=_clear_n, border_color=(0.94, 0.78, 0.78, 1), size_dp=46, radius=12))
        btns.add_widget(self._ui_card_btn('OK', (0.0, 0.62, 0.25, 1), icon='check-circle-outline', size_hint_x=0.65, height=46, elevation=1, on_release=_save_n))
        content.add_widget(btns)

        ndlg = MDDialog(title='Ajouter une note', type='custom', content_cls=content, size_hint=(0.92, None), radius=[18, 18, 18, 18], buttons=[])
        self._fit_dialog(ndlg)
        ndlg.open()
        Clock.schedule_once(lambda dt: setattr(nf, 'focus', True), 0.25)

    def submit_simple_payment(self, x):
        if self.is_transaction_in_progress:
            return
        amt_dec = FinancialEngine.quantize_money(self.txt_simple_amount.text or '0')
        if amt_dec == 0 or abs(amt_dec) >= AppConstants.MAX_AMOUNT_LIMIT:
            self.notify('Montant invalide ou hors limites.', 'error')
            return

        self.is_transaction_in_progress = True
        if getattr(self, 'simple_pay_dialog', None):
            self.simple_pay_dialog.dismiss()

        is_client = self.current_mode in ('client_payment', 'client_pay')
        doc_type = AppConstants.DOC_CLIENT_PAY if is_client else AppConstants.DOC_SUPPLIER_PAY
        label = ('Versement' if is_client else 'Règlement') if amt_dec > 0 else ('Crédit Initial' if is_client else 'Dette Initiale')
        note = self.temp_note.strip() if getattr(self, 'temp_note', '').strip() else label

        try:
            saved_id = self.db.save_transaction({
                'id': self.editing_tx_id,
                'doc_type': doc_type,
                'entity_id': self.selected_entity['id'],
                'amount': float(amt_dec),
                'doc_number': label,
                'note': note,
                'user_name': self.current_user_name,
                'timestamp': str(self.current_editing_date).split('.')[0] if (self.editing_tx_id and self.current_editing_date) else datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'is_simple_payment': True,
            })
            self.notify(f'{label} enregistré avec succès.', 'success')
            if self.db.get_bool_setting('printer_auto', True) and self.db.get_setting('printer_mac', ''):
                det = self.db.get_operation_details(saved_id, is_transfer=False)
                if det:
                    threading.Thread(target=self.print_ticket_bluetooth, args=(det,), daemon=True).start()
            self.temp_note = ''
            self.editing_tx_id = None
            self.current_editing_date = None
            self.load_more_entities(reset=True)
            self.check_and_load_stats()
            if getattr(self, 'pending_dialog', None) and self.pending_dialog.parent:
                self.history_page_offset = 0
                self.history_has_more = True
                self.is_loading_history = False
                if hasattr(self, 'rv_history'):
                    self.rv_history.data = []
                    self.rv_history.scroll_y = 1.0
                self.load_more_history()
            if getattr(self, 'entity_hist_dialog', None) and self.entity_hist_dialog.parent:
                if getattr(self, 'history_target_entity', None) and self.history_target_entity.get('id'):
                    fresh_ent = self.db.get_entity_by_id(self.history_target_entity['id'])
                    if fresh_ent:
                        self.history_target_entity = fresh_ent
                self._reload_entity_history_ui()
        except Exception as exc:
            self.notify(str(exc), 'error')
        finally:
            self.is_transaction_in_progress = False

    def open_entity_history_dialog(self, entity: dict):
        self.history_target_entity = entity
        self.current_eh_filter = AppConstants.FAM_ALL
        today_str = datetime.now().strftime('%Y-%m-%d')
        self.eh_start_date = today_str
        self.eh_end_date = today_str

        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, height=min(dp(605), Window.height * 0.82), padding=[0, dp(2), 0, 0])
        filter_card = MDCard(orientation='vertical', size_hint=(1, None), adaptive_height=True, radius=[12], elevation=0, md_bg_color=(0.96, 0.97, 0.99, 1), line_color=(0.88, 0.9, 0.94, 1), line_width=1, padding=dp(8), spacing=dp(8))

        tabs_container, self.eh_period_btns = self._ui_period_tabs(self.filter_entity_history_list, self._open_entity_hist_date)
        filter_card.add_widget(tabs_container)

        self.btn_eh_filter = MDCard(orientation='horizontal', size_hint=(1, None), height=dp(38), radius=[10], elevation=0, md_bg_color=(0.22, 0.24, 0.28, 1), ripple_behavior=True, padding=[dp(12), 0], spacing=dp(8), on_release=self.open_entity_history_filter_menu)
        self.btn_eh_filter.add_widget(MDIcon(icon='filter-variant', theme_text_color='Custom', text_color=(1, 1, 1, 1), font_size='18sp', size_hint_x=None, width=dp(20), pos_hint={'center_y': 0.5}))
        self.lbl_eh_filter = MDLabel(text='FILTRER PAR TYPE', halign='center', valign='middle', bold=True, font_size='12.5sp', max_lines=1, shorten=True, theme_text_color='Custom', text_color=(1, 1, 1, 1))
        self.btn_eh_filter.add_widget(self.lbl_eh_filter)
        self.btn_eh_filter.add_widget(MDIcon(icon='chevron-down', theme_text_color='Custom', text_color=(1, 1, 1, 1), font_size='20sp', size_hint_x=None, width=dp(20), pos_hint={'center_y': 0.5}))
        filter_card.add_widget(self.btn_eh_filter)

        self.lbl_eh_period = MDLabel(text='', halign='center', font_style='Caption', bold=True, theme_text_color='Secondary', size_hint_y=None, height=dp(18))
        filter_card.add_widget(self.lbl_eh_period)
        content.add_widget(filter_card)

        self.rv_entity_history = HistoryRecycleView()
        self.rv_entity_history.size_hint_y = 1
        content.add_widget(self.rv_entity_history)
        content.add_widget(self._ui_card_btn('FERMER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), height=46, on_release=lambda *a: self.entity_hist_dialog.dismiss() if getattr(self, 'entity_hist_dialog', None) else None))

        self.entity_hist_dialog = MDDialog(
            title=self.fix_text(f"Historique : {entity['name']}"),
            type='custom',
            content_cls=content,
            size_hint=(0.96, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.entity_hist_dialog)
        self.entity_hist_dialog.open()
        self.filter_entity_history_list('auj')

    def _open_entity_hist_date(self, inst):
        dp_dlg = MDDatePicker(mode='range')
        dp_dlg.bind(on_save=lambda i, val, dr: self.filter_entity_history_list('custom', dr[0], dr[-1]) if dr else self.filter_entity_history_list('custom', val, val))
        dp_dlg.open()

    def open_entity_history_filter_menu(self, instance):
        ent = getattr(self, 'history_target_entity', None) or {}
        is_supplier = ent.get('entity_type') == AppConstants.ENTITY_SUPPLIER
        if is_supplier:
            filters = [
                ('Tout', AppConstants.FAM_ALL),
                ("Bon d'Achat (BA)", AppConstants.DOC_BA),
                ('Facture Achat (FF)', AppConstants.DOC_FF),
                ('Commande (DP)', AppConstants.DOC_DP),
                ('Retour Fourn. (RF)', AppConstants.DOC_RF),
                ("Bon d'Entrée (BE)", AppConstants.DOC_BE),
                ('Règlement Fourn.', AppConstants.DOC_SUPPLIER_PAY),
            ]
        elif self.is_seller_mode:
            filters = [
                ('Tout', AppConstants.FAM_ALL),
                ('Bon de Vente (BV)', AppConstants.DOC_BV),
                ('Retour Client (RC)', AppConstants.DOC_RC),
                ('Versement Client', AppConstants.DOC_CLIENT_PAY),
            ]
        else:
            filters = [
                ('Tout', AppConstants.FAM_ALL),
                ('Bon de Vente (BV)', AppConstants.DOC_BV),
                ('Facture (FC)', AppConstants.DOC_FC),
                ('Proforma (FP)', AppConstants.DOC_FP),
                ('Retour Client (RC)', AppConstants.DOC_RC),
                ('Bon de Sortie (BS)', AppConstants.DOC_BS),
                ('Versement Client', AppConstants.DOC_CLIENT_PAY),
            ]

        cur_code = getattr(self, 'current_eh_filter', AppConstants.FAM_ALL)
        options = [
            {
                'label': label,
                'code': code,
                'icon': 'filter-check' if code == AppConstants.FAM_ALL else AppConstants.DOC_VISUALS.get(code, {}).get('icon', 'file'),
                'color': (0.1, 0.45, 0.85, 1) if code == AppConstants.FAM_ALL else AppConstants.DOC_VISUALS.get(code, {}).get('color', (0.4, 0.45, 0.5, 1)),
            }
            for label, code in filters
        ]
        self.hist_filter_dialog = self._ui_filter_menu('Filtrer par type', options, cur_code, self._apply_entity_history_filter)

    def _apply_entity_history_filter(self, code: str):
        self.current_eh_filter = code
        if getattr(self, 'btn_eh_filter', None) and getattr(self, 'lbl_eh_filter', None):
            if code == AppConstants.FAM_ALL:
                self.btn_eh_filter.md_bg_color = (0.22, 0.24, 0.28, 1)
                self.lbl_eh_filter.text = 'FILTRER PAR TYPE'
            else:
                self.btn_eh_filter.md_bg_color = (0, 0.55, 0.8, 1)
                self.lbl_eh_filter.text = f'FILTRE : {code}'
        self._reload_entity_history_ui()

    def filter_entity_history_list(self, ftype: str = 'auj', start_d=None, end_d=None):
        if not hasattr(self, 'rv_entity_history') or not getattr(self, 'history_target_entity', None):
            return
        today = datetime.now().date()
        cal_lbl = 'CAL'
        active_k = ftype
        if ftype == 'custom' and start_d:
            s_d, e_d = start_d, (end_d or start_d)
            cal_lbl = s_d.strftime('%d/%m') if s_d == e_d else 'CAL'
        elif ftype == 'hier':
            s_d = e_d = today - timedelta(days=1)
        elif ftype == 'semaine':
            s_d = today - timedelta(days=(today.weekday() + 2) % 7)
            e_d = today
        elif ftype == 'mois':
            s_d = today.replace(day=1)
            e_d = today
        else:
            active_k = 'auj'
            s_d = e_d = today

        if getattr(self, 'eh_period_btns', None):
            self._activate_period_tab(self.eh_period_btns, active_k, cal_lbl)

        self.eh_start_date = str(s_d)
        self.eh_end_date = str(e_d)
        if getattr(self, 'lbl_eh_period', None):
            self.lbl_eh_period.text = f"Date : {s_d.strftime('%d/%m/%Y')}" if s_d == e_d else f"Du {s_d.strftime('%d/%m/%Y')} au {e_d.strftime('%d/%m/%Y')}"
        self._reload_entity_history_ui()

    def _reload_entity_history_ui(self):
        if not hasattr(self, 'rv_entity_history') or not getattr(self, 'history_target_entity', None):
            return
        cur_f = getattr(self, 'current_eh_filter', AppConstants.FAM_ALL)
        if self.is_seller_mode and cur_f == AppConstants.FAM_ALL:
            cur_f = 'SELLER_ONLY'
        rows = self.db.get_operations_history(
            self.eh_start_date,
            self.eh_end_date,
            entity_id=self.history_target_entity['id'],
            entity_category=self.history_target_entity.get('entity_type'),
            doc_type_filter=cur_f,
            limit=200,
            offset=0,
        )
        formatted = self._format_history_rows(rows)
        if not formatted:
            formatted = [{
                'type_str': 'Information', 'ref_str': '', 'entity_str': 'Aucune opération trouvée.',
                'date_str': '', 'amount_text': '', 'icon': 'information-outline',
                'icon_color': [0.5, 0.5, 0.5, 1], 'bg_color': [1, 1, 1, 1],
                'op_id': None, 'is_transfer': False,
            }]
        self.rv_entity_history.data = formatted
        self.rv_entity_history.scroll_y = 1.0
        self.rv_entity_history.refresh_from_data()

    # --------------------------------------------------------------------------
    # 7.9 DÉPENSES / CHARGES (CLAVIER NUMÉRIQUE + CATÉGORIES ENRICHIES)
    # --------------------------------------------------------------------------
    def open_add_expense_dialog(self, edit_data: dict | None = None):
        if getattr(self, 'expense_dialog', None):
            try:
                self.expense_dialog.dismiss()
            except Exception:
                pass
        Window.release_all_keyboards()

        self.editing_expense_id = edit_data['id'] if edit_data else None
        is_edit = edit_data is not None
        init_cat = edit_data['category'] if is_edit else 'Autre'
        cat_st = AppConstants.get_expense_style(init_cat)
        init_amt = FinancialEngine.format_qty(edit_data['amount']) if is_edit else '0.00'
        self._exp_reset_mode = True

        target_h = min(dp(595), Window.height * 0.90)
        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(8),
            size_hint_y=None,
            height=target_h,
            padding=[0, 0, 0, dp(2)],
        )

        # --- عنوان النافذة ---
        content.add_widget(
            MDLabel(
                text='Modifier la dépense' if is_edit else 'Nouvelle Dépense',
                font_style='H6',
                bold=True,
                theme_text_color='Primary',
                halign='center',
                valign='center',
                size_hint_y=None,
                height=dp(28),
            )
        )

        # --- بطاقة عرض المبلغ (متصلة بلوحة الأرقام مباشرة) ---
        amt_card = MDCard(
            orientation='vertical',
            size_hint=(1, None),
            height=dp(74),
            radius=[14],
            elevation=0,
            md_bg_color=(1, 0.96, 0.96, 1),
            line_color=(0.85, 0.18, 0.18, 1),
            line_width=1.4,
            padding=[dp(12), dp(6)],
            spacing=dp(2),
        )
        amt_card.add_widget(
            MDLabel(
                text='MONTANT DE LA DÉPENSE (DA)',
                font_style='Caption',
                bold=True,
                halign='center',
                theme_text_color='Custom',
                text_color=(0.78, 0.14, 0.14, 1),
                size_hint_y=None,
                height=dp(16),
            )
        )
        amt_row = MDFloatLayout(size_hint_y=1)
        self.exp_amount_field = MDLabel(
            text=init_amt,
            font_size='30sp',
            bold=True,
            halign='center',
            valign='middle',
            theme_text_color='Custom',
            text_color=(0.1, 0.12, 0.16, 1) if is_edit else (0.55, 0.58, 0.64, 1),
            pos_hint={'center_x': 0.5, 'center_y': 0.5},
        )

        def _clear_exp_amount(*_):
            self._exp_reset_mode = True
            self.exp_amount_field.text = '0.00'
            self.exp_amount_field.text_color = (0.55, 0.58, 0.64, 1)

        btn_clr_amt = MDIconButton(
            icon='close-circle-outline',
            theme_text_color='Custom',
            text_color=(0.62, 0.42, 0.42, 1),
            icon_size='20sp',
            pos_hint={'right': 1.0, 'center_y': 0.5},
            on_release=_clear_exp_amount,
        )
        amt_row.add_widget(self.exp_amount_field)
        amt_row.add_widget(btn_clr_amt)
        amt_card.add_widget(amt_row)
        content.add_widget(amt_card)

        # --- بطاقة اختيار الفئة ---
        cat_icon = MDIcon(
            icon=cat_st['icon'],
            theme_text_color='Custom',
            text_color=cat_st['color'],
            font_size='20sp',
            size_hint=(None, None),
            size=(dp(22), dp(22)),
            pos_hint={'center_y': 0.5},
        )
        self.exp_cat_field = MDLabel(
            text=init_cat,
            font_style='Subtitle2',
            bold=True,
            theme_text_color='Primary',
            halign='left',
            valign='middle',
            shorten=True,
            shorten_from='right',
            pos_hint={'center_y': 0.5},
        )

        def _set_exp_cat(val: str):
            st = AppConstants.get_expense_style(val)
            self.exp_cat_field.text = val
            cat_icon.icon = st['icon']
            cat_icon.text_color = st['color']
            cat_card.md_bg_color = st['bg']
            cat_card.line_color = (st['color'][0], st['color'][1], st['color'][2], 0.35)

        def _open_cat_picker(*args):
            if hasattr(self, 'exp_desc_field'):
                self.exp_desc_field.focus = False
            Window.release_all_keyboards()
            if getattr(self, 'exp_cat_dialog', None):
                try:
                    self.exp_cat_dialog.dismiss()
                except Exception:
                    pass
            options = [
                {
                    'label': c,
                    'code': c,
                    'icon': AppConstants.get_expense_style(c)['icon'],
                    'color': AppConstants.get_expense_style(c)['color'],
                }
                for c in AppConstants.EXPENSE_CATEGORIES
            ]
            self.exp_cat_dialog = self._ui_filter_menu('Catégorie de dépense', options, self.exp_cat_field.text, _set_exp_cat)

        cat_card = MDCard(
            orientation='horizontal',
            size_hint=(1, None),
            height=dp(46),
            radius=[12],
            elevation=0,
            md_bg_color=cat_st['bg'],
            line_color=(cat_st['color'][0], cat_st['color'][1], cat_st['color'][2], 0.35),
            line_width=1,
            padding=[dp(12), 0, dp(10), 0],
            spacing=dp(10),
            ripple_behavior=True,
            on_release=_open_cat_picker,
        )
        cat_card.add_widget(cat_icon)
        cat_card.add_widget(self.exp_cat_field)
        cat_card.add_widget(
            MDIcon(
                icon='chevron-down',
                theme_text_color='Secondary',
                font_size='21sp',
                size_hint_x=None,
                width=dp(22),
                pos_hint={'center_y': 0.5},
            )
        )
        content.add_widget(cat_card)

        # --- حقل الملاحظة / الوصف ---
        desc_card = MDCard(
            orientation='horizontal',
            size_hint=(1, None),
            height=dp(52),
            radius=[12],
            elevation=0,
            md_bg_color=(0.96, 0.97, 0.99, 1),
            line_color=(0.86, 0.88, 0.93, 1),
            line_width=1,
            padding=[dp(12), dp(2), dp(8), dp(2)],
            spacing=dp(8),
        )
        self.exp_desc_field = SmartTextField(
            hint_text='Description / Note (optionnel)',
            text=edit_data['description'] if is_edit else '',
            mode='line',
            active_line=False,
            line_color_normal=(0, 0, 0, 0),
            line_color_focus=(0, 0, 0, 0),
            icon_right='note-text-outline',
            pos_hint={'center_y': 0.5},
        )
        desc_card.add_widget(self.exp_desc_field)
        content.add_widget(desc_card)

        # --- لوحة الأرقام المدمجة ---
        def _digit(d: str):
            if getattr(self, 'exp_desc_field', None) and self.exp_desc_field.focus:
                self.exp_desc_field.focus = False
                Window.release_all_keyboards()
            cur = self.exp_amount_field.text
            if self._exp_reset_mode:
                self.exp_amount_field.text = '0.' if d == '.' else d
                self._exp_reset_mode = False
            elif d == '.':
                if '.' not in cur and len(cur) < 12:
                    self.exp_amount_field.text = '0.' if not cur else cur + '.'
            elif cur in ('0', '0.00'):
                self.exp_amount_field.text = d
            elif len(cur) < 12:
                self.exp_amount_field.text = cur + d
            self.exp_amount_field.text_color = (0.1, 0.12, 0.16, 1)

        def _bkspc(*_):
            if getattr(self, 'exp_desc_field', None) and self.exp_desc_field.focus:
                self.exp_desc_field.focus = False
                Window.release_all_keyboards()
            cur = self.exp_amount_field.text
            if self._exp_reset_mode and cur == '0.00':
                return
            self._exp_reset_mode = False
            new_val = cur[:-1] if cur else ''
            if not new_val:
                self.exp_amount_field.text = '0.00'
                self.exp_amount_field.text_color = (0.55, 0.58, 0.64, 1)
                self._exp_reset_mode = True
            else:
                self.exp_amount_field.text = new_val
                self.exp_amount_field.text_color = (0.1, 0.12, 0.16, 1)

        grid = MDGridLayout(cols=3, spacing=dp(7), size_hint_y=1)
        for key in ('7', '8', '9', '4', '5', '6', '1', '2', '3', '.', '0', 'DEL'):
            if key == 'DEL':
                del_card = MDCard(
                    md_bg_color=(1, 0.93, 0.93, 1),
                    line_color=(0.94, 0.78, 0.78, 1),
                    line_width=1,
                    radius=[12],
                    elevation=0,
                    ripple_behavior=True,
                    on_release=_bkspc,
                )
                d_fl = MDFloatLayout()
                d_fl.add_widget(
                    MDIcon(
                        icon='backspace-outline',
                        theme_text_color='Custom',
                        text_color=(0.82, 0.14, 0.14, 1),
                        font_size='22sp',
                        size_hint=(None, None),
                        size=(dp(24), dp(24)),
                        halign='center',
                        valign='middle',
                        pos_hint={'center_x': 0.5, 'center_y': 0.5},
                    )
                )
                del_card.add_widget(d_fl)
                grid.add_widget(del_card)
            else:
                key_card = MDCard(
                    md_bg_color=(0.96, 0.97, 0.99, 1),
                    line_color=(0.88, 0.90, 0.94, 1),
                    line_width=1,
                    radius=[12],
                    elevation=0,
                    ripple_behavior=True,
                    on_release=lambda x, k=key: _digit(k),
                )
                key_card.add_widget(
                    MDLabel(
                        text=key,
                        halign='center',
                        valign='middle',
                        font_size='21sp',
                        bold=True,
                        theme_text_color='Custom',
                        text_color=(0.12, 0.14, 0.18, 1),
                    )
                )
                grid.add_widget(key_card)
        content.add_widget(grid)

        # --- زر السجل في الأسفل بعرض كامل فوق زر التسجيل ---
        if not is_edit:
            content.add_widget(
                self._ui_card_btn(
                    "VOIR L'HISTORIQUE",
                    (0.09, 0.44, 0.84, 1),
                    text_color=(1, 1, 1, 1),
                    icon='history',
                    height=42,
                    radius=12,
                    size_hint_x=1,
                    font_size='13.5sp',
                    elevation=2,
                    on_release=lambda x: [
                        Window.release_all_keyboards(),
                        self.expense_dialog.dismiss(),
                        self.open_expense_history_dialog(),
                    ],
                )
            )

        def _on_save_tap(*args):
            if hasattr(self, 'exp_desc_field'):
                self.exp_desc_field.focus = False
            Window.release_all_keyboards()
            self.submit_expense(None)

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(48))
        btns.add_widget(
            self._ui_card_btn(
                'ANNULER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                size_hint_x=0.36,
                height=48,
                font_size='13sp',
                on_release=lambda x: [Window.release_all_keyboards(), self.expense_dialog.dismiss()],
            )
        )
        btns.add_widget(
            self._ui_card_btn(
                'ENREGISTRER',
                (0.0, 0.62, 0.25, 1),
                icon='check-circle-outline',
                size_hint_x=0.64,
                height=48,
                font_size='14sp',
                elevation=1,
                on_release=_on_save_tap,
            )
        )
        content.add_widget(btns)

        self.expense_dialog = MDDialog(
            title='',
            type='custom',
            content_cls=content,
            size_hint=(0.94, None),
            radius=[20, 20, 20, 20],
            buttons=[],
        )
        self._fit_dialog(self.expense_dialog, top_pad=14, side_pad=14, bottom_pad=12, spacer_top=0)
        self.expense_dialog.open()

    def submit_expense(self, instance):
        try:
            self.db.save_expense(
                description=self.exp_desc_field.get_value().strip() or '-',
                amount=self.exp_amount_field.text,
                category=self.exp_cat_field.text.strip() or 'Autre',
                recorded_by=self.current_user_name,
                expense_id=getattr(self, 'editing_expense_id', None),
            )
            if getattr(self, 'expense_dialog', None):
                self.expense_dialog.dismiss()
            self.notify('Dépense enregistrée avec succès.', 'success')
            self.check_and_load_stats()
            if self.editing_expense_id:
                self.open_expense_history_dialog()
        except Exception as exc:
            self.notify(str(exc), 'error')

    def open_expense_history_dialog(self):
        if getattr(self, 'history_exp_dialog', None):
            try:
                self.history_exp_dialog.dismiss()
            except Exception:
                pass
        today_str = datetime.now().strftime('%Y-%m-%d')
        self.exp_start_date = today_str
        self.exp_end_date = today_str
        self.current_exp_category = AppConstants.FAM_ALL

        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, height=min(dp(595), Window.height * 0.82), padding=[0, dp(2), 0, 0])
        filter_card = MDCard(orientation='vertical', size_hint=(1, None), adaptive_height=True, radius=[12], elevation=0, md_bg_color=(0.96, 0.97, 0.99, 1), line_color=(0.88, 0.9, 0.94, 1), line_width=1, padding=dp(8), spacing=dp(8))

        tabs_container, self.exp_period_btns = self._ui_period_tabs(self.filter_expenses_list, self.open_expense_date_picker)
        filter_card.add_widget(tabs_container)

        self.btn_exp_filter = MDCard(orientation='horizontal', size_hint=(1, None), height=dp(38), radius=[10], elevation=0, md_bg_color=(0.22, 0.24, 0.28, 1), ripple_behavior=True, padding=[dp(12), 0], spacing=dp(8), on_release=self.open_expense_filter_menu)
        self.btn_exp_filter.add_widget(MDIcon(icon='filter-variant', theme_text_color='Custom', text_color=(1, 1, 1, 1), font_size='18sp', size_hint_x=None, width=dp(20), pos_hint={'center_y': 0.5}))
        self.lbl_exp_filter = MDLabel(text='FILTRER PAR CATÉGORIE', halign='center', valign='middle', bold=True, font_size='13sp', theme_text_color='Custom', text_color=(1, 1, 1, 1))
        self.btn_exp_filter.add_widget(self.lbl_exp_filter)
        self.btn_exp_filter.add_widget(MDIcon(icon='chevron-down', theme_text_color='Custom', text_color=(1, 1, 1, 1), font_size='20sp', size_hint_x=None, width=dp(20), pos_hint={'center_y': 0.5}))
        filter_card.add_widget(self.btn_exp_filter)

        self.lbl_exp_period = MDLabel(text='', halign='center', font_style='Caption', bold=True, theme_text_color='Secondary', size_hint_y=None, height=dp(18))
        filter_card.add_widget(self.lbl_exp_period)
        content.add_widget(filter_card)

        scroll = MDScrollView(size_hint_y=1, bar_width=dp(3))
        self.exp_list_layout = MDBoxLayout(orientation='vertical', spacing=dp(8), adaptive_height=True, padding=[0, dp(2), dp(2), dp(4)])
        scroll.add_widget(self.exp_list_layout)
        content.add_widget(scroll)

        total_card = MDCard(orientation='horizontal', size_hint=(1, None), height=dp(46), radius=[12], elevation=0, md_bg_color=(1, 0.95, 0.95, 1), line_color=(0.92, 0.72, 0.72, 1), line_width=1, padding=[dp(14), 0])
        total_card.add_widget(MDLabel(text='TOTAL DÉPENSES', font_style='Caption', bold=True, theme_text_color='Custom', text_color=(0.75, 0.15, 0.15, 1), valign='middle', size_hint_x=0.45))
        self.lbl_exp_total = MDLabel(text='0,00 DA', bold=True, font_style='H6', theme_text_color='Custom', text_color=(0.82, 0.1, 0.1, 1), halign='right', valign='middle', size_hint_x=0.55)
        total_card.add_widget(self.lbl_exp_total)
        content.add_widget(total_card)

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(48))
        btns.add_widget(self._ui_card_btn('FERMER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.38, on_release=lambda x: self.history_exp_dialog.dismiss()))
        btns.add_widget(self._ui_card_btn('NOUVELLE', self.theme_cls.primary_color, icon='plus-circle-outline', size_hint_x=0.62, font_size='14sp', elevation=1, on_release=lambda x: [self.history_exp_dialog.dismiss(), self.open_add_expense_dialog()]))
        content.add_widget(btns)

        self.history_exp_dialog = MDDialog(
            title='Historique des dépenses',
            type='custom',
            content_cls=content,
            size_hint=(0.96, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.history_exp_dialog)
        self.history_exp_dialog.open()
        self.filter_expenses_list('auj')

    def open_expense_date_picker(self, instance):
        dp_dlg = MDDatePicker(mode='range')
        dp_dlg.bind(on_save=lambda inst, val, dr: self.filter_expenses_list('custom', dr[0], dr[-1]) if dr else self.filter_expenses_list('custom', val, val))
        dp_dlg.open()

    def filter_expenses_list(self, ftype: str = 'auj', start_d=None, end_d=None):
        today = datetime.now().date()
        cal_lbl = 'CAL'
        active_k = ftype
        if ftype == 'custom' and start_d:
            s_d, e_d = start_d, (end_d or start_d)
            cal_lbl = s_d.strftime('%d/%m') if s_d == e_d else 'CAL'
        elif ftype == 'hier':
            s_d = e_d = today - timedelta(days=1)
        elif ftype == 'semaine':
            s_d = today - timedelta(days=(today.weekday() + 2) % 7)
            e_d = today
        elif ftype == 'mois':
            s_d = today.replace(day=1)
            e_d = today
        else:
            active_k = 'auj'
            s_d = e_d = today

        if getattr(self, 'exp_period_btns', None):
            self._activate_period_tab(self.exp_period_btns, active_k, cal_lbl)

        self.exp_start_date = str(s_d)
        self.exp_end_date = str(e_d)
        if getattr(self, 'lbl_exp_period', None):
            self.lbl_exp_period.text = f"Date : {s_d.strftime('%d/%m/%Y')}" if s_d == e_d else f"Du {s_d.strftime('%d/%m/%Y')} au {e_d.strftime('%d/%m/%Y')}"
        self._load_expenses_ui()

    def open_expense_filter_menu(self, instance):
        cur_cat = getattr(self, 'current_exp_category', AppConstants.FAM_ALL)
        options = [{'label': 'Tout', 'code': AppConstants.FAM_ALL, 'icon': 'filter-check', 'color': (0.1, 0.45, 0.85, 1)}]
        for c in AppConstants.EXPENSE_CATEGORIES:
            st = AppConstants.get_expense_style(c)
            options.append({'label': c, 'code': c, 'icon': st['icon'], 'color': st['color']})
        self.hist_filter_dialog = self._ui_filter_menu('Filtrer par catégorie', options, cur_cat, self._apply_expense_filter)

    def _apply_expense_filter(self, code: str):
        self.current_exp_category = code
        if getattr(self, 'btn_exp_filter', None) and getattr(self, 'lbl_exp_filter', None):
            if code == AppConstants.FAM_ALL:
                self.btn_exp_filter.md_bg_color = (0.22, 0.24, 0.28, 1)
                self.lbl_exp_filter.text = 'FILTRER PAR CATÉGORIE'
            else:
                self.btn_exp_filter.md_bg_color = (0, 0.55, 0.8, 1)
                self.lbl_exp_filter.text = f'FILTRE : {code}'
        self._load_expenses_ui()

    def _load_expenses_ui(self):
        self.exp_list_layout.clear_widgets()
        rows = self.db.get_expenses(self.exp_start_date, self.exp_end_date)
        cat_filter = getattr(self, 'current_exp_category', AppConstants.FAM_ALL)
        if cat_filter != AppConstants.FAM_ALL:
            rows = [r for r in rows if r.get('category') == cat_filter]

        total = Decimal('0.00')
        for item in rows:
            amt = FinancialEngine.to_decimal(item.get('amount', 0))
            total += amt
            st = AppConstants.get_expense_style(item.get('category', 'Autre'))
            card = MDCard(
                orientation='horizontal',
                padding=[dp(10), dp(10)],
                spacing=dp(10),
                size_hint_y=None,
                height=dp(84),
                radius=[12],
                md_bg_color=(1, 1, 1, 1),
                line_color=(0.89, 0.91, 0.94, 1),
                line_width=1,
                elevation=0,
            )
            card.add_widget(self._ui_icon_pill(st['icon'], st['bg'], st['color'], size_dp=40, radius=20, icon_size='22sp'))

            info = MDBoxLayout(orientation='vertical', size_hint_x=1, spacing=dp(2), pos_hint={'center_y': 0.5})
            info.add_widget(MDLabel(text=f"{item['category']} • {FinancialEngine.format_money(amt)}", font_style='Subtitle2', bold=True, theme_text_color='Primary', shorten=True))
            info.add_widget(MDLabel(text=self.fix_text(item['description']), font_style='Caption', theme_text_color='Secondary', max_lines=2, shorten=False))
            info.add_widget(MDLabel(text=str(item['expense_date'])[:16], font_style='Caption', theme_text_color='Hint'))
            card.add_widget(info)

            acts = MDBoxLayout(orientation='horizontal', size_hint=(None, None), size=(dp(74), dp(34)), spacing=dp(6), pos_hint={'center_y': 0.5})
            acts.add_widget(self._ui_icon_pill('pencil-outline', (0.91, 0.95, 1, 1), (0.08, 0.45, 0.82, 1), on_release=lambda x, i=item: [self.history_exp_dialog.dismiss(), self.open_add_expense_dialog(edit_data=i)], size_dp=34, icon_size='18sp'))
            acts.add_widget(self._ui_icon_pill('trash-can-outline', (1, 0.93, 0.93, 1), (0.82, 0.12, 0.12, 1), on_release=lambda x, eid=item['id']: self.confirm_delete_expense(eid), size_dp=34, icon_size='18sp'))
            card.add_widget(acts)
            self.exp_list_layout.add_widget(card)

        self.lbl_exp_total.text = FinancialEngine.format_money(total)
        if not rows:
            empty_card = MDCard(orientation='vertical', size_hint=(1, None), height=dp(140), radius=[12], elevation=0, md_bg_color=(0.98, 0.98, 0.99, 1), line_color=(0.90, 0.91, 0.94, 1), line_width=1, padding=dp(16), spacing=dp(8))
            empty_card.add_widget(MDBoxLayout())
            empty_card.add_widget(MDIcon(icon='receipt-text-outline', theme_text_color='Hint', font_size='38sp', halign='center', pos_hint={'center_x': 0.5}))
            empty_card.add_widget(MDLabel(text='Aucune dépense sur cette période.', halign='center', font_style='Body2', theme_text_color='Hint', adaptive_height=True))
            empty_card.add_widget(MDBoxLayout())
            self.exp_list_layout.add_widget(empty_card)

    def confirm_delete_expense(self, exp_id: int):
        def _do_del():
            self.db.delete_expense(exp_id)
            self.notify('Dépense supprimée.', 'success')
            self._load_expenses_ui()
            self.check_and_load_stats()

        self.del_exp_dialog = self._ui_confirm_dialog(
            title='Confirmer la suppression',
            message='Voulez-vous vraiment supprimer cette dépense ?',
            confirm_text='SUPPRIMER',
            icon='trash-can-outline',
            confirm_icon='trash-can-outline',
            on_confirm=_do_del,
        )

    # --------------------------------------------------------------------------
    # 7.10 HISTORIQUE GLOBAL, DÉTAILS, MODIFICATION (Y COMPRIS BI) & PDF INTELLIGENT
    # --------------------------------------------------------------------------
    def show_pending_dialog(self):
        self.current_history_filter = AppConstants.FAM_ALL
        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, height=min(dp(605), Window.height * 0.82), padding=[0, dp(2), 0, 0])
        filter_card = MDCard(orientation='vertical', size_hint=(1, None), adaptive_height=True, radius=[12], elevation=0, md_bg_color=(0.96, 0.97, 0.99, 1), line_color=(0.88, 0.9, 0.94, 1), line_width=1, padding=dp(8), spacing=dp(8))

        tabs_container, self.hist_period_btns = self._ui_period_tabs(self.filter_history_list, self.open_history_date_picker)
        filter_card.add_widget(tabs_container)

        if not self.is_seller_mode:
            self.btn_hist_filter = MDCard(orientation='horizontal', size_hint=(1, None), height=dp(38), radius=[10], elevation=0, md_bg_color=(0.22, 0.24, 0.28, 1), ripple_behavior=True, padding=[dp(12), 0], spacing=dp(8), on_release=self.open_history_filter_menu)
            self.btn_hist_filter.add_widget(MDIcon(icon='filter-variant', theme_text_color='Custom', text_color=(1, 1, 1, 1), font_size='18sp', size_hint_x=None, width=dp(20), pos_hint={'center_y': 0.5}))
            self.lbl_hist_filter = MDLabel(text='FILTRER PAR TYPE', halign='center', valign='middle', bold=True, font_size='12.5sp', max_lines=1, shorten=True, theme_text_color='Custom', text_color=(1, 1, 1, 1))
            self.btn_hist_filter.add_widget(self.lbl_hist_filter)
            self.btn_hist_filter.add_widget(MDIcon(icon='chevron-down', theme_text_color='Custom', text_color=(1, 1, 1, 1), font_size='20sp', size_hint_x=None, width=dp(20), pos_hint={'center_y': 0.5}))
            filter_card.add_widget(self.btn_hist_filter)

        self.lbl_hist_period = MDLabel(text='', halign='center', font_style='Caption', bold=True, theme_text_color='Secondary', size_hint_y=None, height=dp(18))
        filter_card.add_widget(self.lbl_hist_period)
        content.add_widget(filter_card)

        self.rv_history = HistoryRecycleView()
        self.rv_history.size_hint_y = 1
        content.add_widget(self.rv_history)
        content.add_widget(self._ui_card_btn('FERMER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), height=46, on_release=lambda *a: self.pending_dialog.dismiss() if getattr(self, 'pending_dialog', None) else None))

        self.pending_dialog = MDDialog(
            title='Historique des Opérations',
            type='custom',
            content_cls=content,
            size_hint=(0.96, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.pending_dialog)
        self.pending_dialog.open()
        self.filter_history_list('auj')

    def open_history_filter_menu(self, instance):
        filters = [
            ('Tout', AppConstants.FAM_ALL),
            ('Bon de Vente (BV)', AppConstants.DOC_BV),
            ("Bon d'Achat (BA)", AppConstants.DOC_BA),
            ('Facture (FC)', AppConstants.DOC_FC),
            ('Facture Achat (FF)', AppConstants.DOC_FF),
            ('Proforma (FP)', AppConstants.DOC_FP),
            ('Commande (DP)', AppConstants.DOC_DP),
            ('Retour Client (RC)', AppConstants.DOC_RC),
            ('Retour Fourn. (RF)', AppConstants.DOC_RF),
            ("Bon d'Entrée (BE)", AppConstants.DOC_BE),
            ('Bon de Sortie (BS)', AppConstants.DOC_BS),
            ('Transfert (TR)', AppConstants.DOC_TR),
            ('Bon Initial (BI)', AppConstants.DOC_BI),
            ('Versement Client', AppConstants.DOC_CLIENT_PAY),
            ('Règlement Fourn.', AppConstants.DOC_SUPPLIER_PAY),
        ]
        cur_code = getattr(self, 'current_history_filter', AppConstants.FAM_ALL)
        options = [
            {
                'label': label,
                'code': code,
                'icon': 'filter-check' if code == AppConstants.FAM_ALL else AppConstants.DOC_VISUALS.get(code, {}).get('icon', 'file'),
                'color': (0.1, 0.45, 0.85, 1) if code == AppConstants.FAM_ALL else AppConstants.DOC_VISUALS.get(code, {}).get('color', (0.4, 0.45, 0.5, 1)),
            }
            for label, code in filters
        ]
        self.hist_filter_dialog = self._ui_filter_menu('Filtrer par type', options, cur_code, self._apply_history_filter)

    def _apply_history_filter(self, code: str):
        self.current_history_filter = code
        if getattr(self, 'btn_hist_filter', None) and getattr(self, 'lbl_hist_filter', None):
            if code == AppConstants.FAM_ALL:
                self.btn_hist_filter.md_bg_color = (0.22, 0.24, 0.28, 1)
                self.lbl_hist_filter.text = 'FILTRER PAR TYPE'
            else:
                self.btn_hist_filter.md_bg_color = (0, 0.55, 0.8, 1)
                self.lbl_hist_filter.text = f'FILTRE : {code}'
        self.history_page_offset = 0
        self.history_has_more = True
        self.is_loading_history = False
        if hasattr(self, 'rv_history'):
            self.rv_history.data = []
            self.rv_history.scroll_y = 1.0
        self.load_more_history()

    def open_history_date_picker(self, instance):
        dp_dlg = MDDatePicker(mode='range')
        dp_dlg.bind(on_save=lambda inst, val, dr: self.filter_history_list('custom', dr[0], dr[-1]) if dr else self.filter_history_list('custom', val, val))
        dp_dlg.open()

    def filter_history_list(self, ftype: str = 'auj', start_d=None, end_d=None):
        today = datetime.now().date()
        cal_lbl = 'CAL'
        active_k = ftype
        if ftype == 'custom' and start_d:
            self.history_start_date = start_d
            self.history_end_date = end_d or start_d
            cal_lbl = start_d.strftime('%d/%m') if start_d == self.history_end_date else 'CAL'
        elif ftype == 'hier':
            self.history_start_date = self.history_end_date = today - timedelta(days=1)
        elif ftype == 'semaine':
            self.history_start_date = today - timedelta(days=(today.weekday() + 2) % 7)
            self.history_end_date = today
        elif ftype == 'mois':
            self.history_start_date = today.replace(day=1)
            self.history_end_date = today
        else:
            active_k = 'auj'
            self.history_start_date = self.history_end_date = today

        if getattr(self, 'hist_period_btns', None):
            self._activate_period_tab(self.hist_period_btns, active_k, cal_lbl)

        if getattr(self, 'lbl_hist_period', None):
            if self.history_start_date == self.history_end_date:
                self.lbl_hist_period.text = f"Date : {self.history_start_date.strftime('%d/%m/%Y')}"
            else:
                self.lbl_hist_period.text = f"Du {self.history_start_date.strftime('%d/%m/%Y')} au {self.history_end_date.strftime('%d/%m/%Y')}"

        self.history_page_offset = 0
        self.history_has_more = True
        self.is_loading_history = False
        if hasattr(self, 'rv_history'):
            self.rv_history.data = []
            self.rv_history.scroll_y = 1.0
            self.rv_history.loading_lock = False
        self.load_more_history()

    def load_more_history(self):
        if self.is_loading_history or not getattr(self, 'history_has_more', True):
            if hasattr(self, 'rv_history'):
                self.rv_history.loading_lock = False
            return
        self.is_loading_history = True
        self._history_fetch_seq += 1
        seq = self._history_fetch_seq
        offset_val = self.history_page_offset
        s_date = str(self.history_start_date)
        e_date = str(self.history_end_date)
        doc_filter = 'SELLER_ONLY' if self.is_seller_mode else self.current_history_filter
        threading.Thread(
            target=self._load_history_worker,
            args=(s_date, e_date, doc_filter, offset_val, seq),
            daemon=True,
        ).start()

    def _load_history_worker(self, s_date: str, e_date: str, doc_filter: str, offset_val: int, seq: int):
        try:
            rows = self.db.get_operations_history(
                s_date,
                e_date,
                doc_type_filter=doc_filter,
                limit=self.history_batch_size,
                offset=offset_val,
            )
            if seq != self._history_fetch_seq:
                return
            formatted = self._format_history_rows(rows)
            Clock.schedule_once(lambda dt: self._append_history_to_rv(formatted, len(rows), offset_val == 0, seq), 0)
        except Exception:
            Clock.schedule_once(lambda dt: self._finish_history_loading(), 0)

    @mainthread
    def _append_history_to_rv(self, formatted: list, db_count: int, is_reset: bool, seq: int):
        if seq != self._history_fetch_seq:
            return
        self.history_has_more = db_count == self.history_batch_size
        if is_reset and not formatted:
            formatted = [{
                'type_str': 'Information', 'ref_str': '', 'entity_str': 'Aucune opération trouvée.',
                'date_str': '', 'amount_text': '', 'icon': 'information-outline',
                'icon_color': [0.5, 0.5, 0.5, 1], 'bg_color': [1, 1, 1, 1],
                'op_id': None, 'is_transfer': False,
            }]
        if hasattr(self, 'rv_history') and self.rv_history:
            if is_reset:
                self.history_page_offset = db_count
            else:
                self.history_page_offset += db_count
            self._preserve_rv_scroll_on_append(self.rv_history, formatted, is_reset)
        self.is_loading_history = False

    @mainthread
    def _finish_history_loading(self):
        self.is_loading_history = False
        if hasattr(self, 'rv_history') and self.rv_history:
            self.rv_history.loading_lock = False

    def _format_history_rows(self, rows: list) -> list:
        out = []
        for r in rows:
            dt_code = r['doc_type']
            vis = AppConstants.DOC_VISUALS.get(dt_code, {'name': dt_code, 'icon': 'file', 'color': (0.2, 0.2, 0.2, 1), 'bg': (1, 1, 1, 1)})
            doc_name = vis['name']
            icon = vis['icon']
            icol = vis['color']
            bg_c = vis['bg']

            amt = FinancialEngine.to_decimal(r.get('total_amount', 0))
            if r.get('is_transfer'):
                src = AppConstants.LOC_LABELS.get(r.get('source_location'), 'Magasin')
                dst = AppConstants.LOC_LABELS.get(r.get('destination_location'), 'Dépôt')
                ent_str = f'{src} -> {dst}'
                amt_txt = 'Stock'
            else:
                ent_str = r.get('entity_name') or AppConstants.DEFAULT_CLIENT_NAME
                if dt_code in AppConstants.QTY_ONLY_DOC_TYPES:
                    amt_txt = 'Stock / Qté'
                elif dt_code in (AppConstants.DOC_CLIENT_PAY, AppConstants.DOC_SUPPLIER_PAY):
                    if amt < 0:
                        doc_name = 'Crédit Initial' if dt_code == AppConstants.DOC_CLIENT_PAY else 'Dette Initiale'
                        icon, icol, bg_c = 'cash-minus', (0.85, 0, 0, 1), (1, 0.94, 0.94, 1)
                        amt_txt = f'- {FinancialEngine.format_money(abs(amt))}'
                    else:
                        amt_txt = f'+ {FinancialEngine.format_money(amt)}'
                else:
                    amt_txt = FinancialEngine.format_money(abs(amt))

            d_str = str(r.get('date', ''))[:16]
            out.append({
                'type_str': doc_name,
                'ref_str': f"{r.get('doc_number', '')} • {r.get('user_name', AppConstants.ADMIN_USERNAME)}",
                'entity_str': ent_str,
                'date_str': d_str,
                'amount_text': amt_txt,
                'icon': icon,
                'icon_color': icol,
                'bg_color': bg_c,
                'op_id': r['id'],
                'is_transfer': bool(r.get('is_transfer', False)),
            })
        return out

    def view_transaction_details(self, op_id: int, is_transfer: bool = False):
        details = self.db.get_operation_details(op_id, is_transfer=is_transfer)
        if not details:
            self.notify('Opération introuvable.', 'error')
            return
        tx = details['transaction']
        ent = details['entity']
        items = details['items']
        dt_code = tx['doc_type']
        vis = AppConstants.DOC_VISUALS.get(
            dt_code,
            {'name': dt_code, 'icon': 'file-document', 'color': (0.2, 0.3, 0.5, 1), 'bg': (0.96, 0.97, 0.99, 1)},
        )

        ent_cat = tx.get('entity_category', AppConstants.get_entity_type(dt_code))
        is_supp = ent_cat == AppConstants.ENTITY_SUPPLIER
        is_comptoir = AppConstants.is_default_entity(tx.get('entity_id'), ent_cat)
        total_amt = FinancialEngine.to_decimal(tx.get('total_amount', 0))
        paid_amt = FinancialEngine.to_decimal(tx.get('paid_amount', 0))
        timbre_val = FinancialEngine.to_decimal(tx.get('stamp_duty', 0))

        doc_title = vis['name']
        doc_icon = vis.get('icon', 'file-document')
        doc_col = vis.get('color', (0.2, 0.3, 0.5, 1))
        doc_bg = vis.get('bg', (0.96, 0.97, 0.99, 1))
        if dt_code in (AppConstants.DOC_CLIENT_PAY, AppConstants.DOC_SUPPLIER_PAY) and total_amt < 0:
            doc_title = 'Dette Initiale' if dt_code == AppConstants.DOC_SUPPLIER_PAY else 'Crédit Initial'
            doc_icon = 'cash-minus'
            doc_col = (0.85, 0.1, 0.1, 1)
            doc_bg = (1, 0.94, 0.94, 1)

        is_today = str(tx.get('date', '')).split(' ')[0] == datetime.now().strftime('%Y-%m-%d')
        can_edit = (not self.is_seller_mode) or (is_today and dt_code != AppConstants.DOC_BI)
        can_del = (not self.is_seller_mode) or is_today

        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, height=min(dp(610), Window.height * 0.85), padding=[0, 0, 0, 0])

        top_bar = MDBoxLayout(orientation='horizontal', size_hint=(1, None), height=dp(38), spacing=dp(8))
        top_bar.add_widget(MDLabel(text='Détails', font_style='H6', bold=True, theme_text_color='Primary', valign='center'))
        if can_del:
            top_bar.add_widget(
                self._ui_icon_pill(
                    'trash-can-outline',
                    (1, 0.93, 0.93, 1),
                    (0.82, 0.12, 0.12, 1),
                    on_release=lambda x: self.confirm_delete_transaction(op_id, is_transfer),
                    border_color=(0.94, 0.74, 0.74, 1),
                    size_dp=36,
                    icon_size='19sp',
                )
            )
        content.add_widget(top_bar)

        hbox = MDCard(orientation='vertical', size_hint=(1, None), adaptive_height=True, padding=dp(12), spacing=dp(6), md_bg_color=doc_bg, line_color=(doc_col[0], doc_col[1], doc_col[2], 0.25), line_width=1, radius=[14], elevation=0)

        top_info = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(42), spacing=dp(10))
        top_info.add_widget(self._ui_icon_pill(doc_icon, (1, 1, 1, 0.9), doc_col, size_dp=40, radius=20, icon_size='22sp'))

        title_col = MDBoxLayout(orientation='vertical', pos_hint={'center_y': 0.5}, spacing=dp(2))
        title_col.add_widget(MDLabel(text=self.fix_text(doc_title), bold=True, font_style='Subtitle1', theme_text_color='Primary', adaptive_height=True))
        ref_str = tx.get('doc_number', '')
        date_str = str(tx.get('date', ''))[:16]
        sub_txt = f'{ref_str}  •  {date_str}' if ref_str else date_str
        title_col.add_widget(MDLabel(text=self.fix_text(sub_txt), font_style='Caption', theme_text_color='Secondary', adaptive_height=True))
        top_info.add_widget(title_col)
        hbox.add_widget(top_info)
        hbox.add_widget(MDSeparator(height=dp(1), color=(0, 0, 0, 0.08)))

        if is_transfer:
            src = AppConstants.LOC_LABELS.get(tx.get('source_location'), 'Magasin')
            dst = AppConstants.LOC_LABELS.get(tx.get('destination_location'), 'Dépôt')
            hbox.add_widget(self._ui_wrap_label(f'Trajet : {src}  ->  {dst}', font_style='Subtitle2', bold=True, theme_color='Primary', min_h=20))
        else:
            lbl_e = 'Fournisseur' if is_supp else 'Client'
            hbox.add_widget(self._ui_wrap_label(f"{lbl_e} : {self.fix_text(ent.get('name', AppConstants.DEFAULT_CLIENT_NAME))}", font_style='Subtitle2', bold=True, theme_color='Primary', min_h=20))

        if not is_transfer and dt_code not in AppConstants.QTY_ONLY_DOC_TYPES:
            if timbre_val > 0:
                hbox.add_widget(MDLabel(text=f'Droit de Timbre : {FinancialEngine.format_money(timbre_val)}', theme_text_color='Custom', text_color=(0.5, 0, 0.5, 1), font_style='Caption', bold=True, adaptive_height=True))
            hbox.add_widget(MDLabel(text=f'Montant : {FinancialEngine.format_money(abs(total_amt))}', bold=True, font_style='H5', theme_text_color='Custom', text_color=doc_col, adaptive_height=True))
            if AppConstants.FINANCIAL_FACTORS.get(dt_code, 0) == 1 and not is_comptoir:
                rem = FinancialEngine.quantize_money(total_amt - paid_amt)
                if abs(rem) < Decimal('0.05'):
                    hbox.add_widget(MDLabel(text='Statut : Payée', theme_text_color='Custom', text_color=(0, 0.62, 0.2, 1), font_style='Subtitle2', bold=True, adaptive_height=True))
                else:
                    pay_row = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(22), spacing=dp(8))
                    pay_row.add_widget(MDLabel(text=f"{'Règlement' if is_supp else 'Versement'} : {FinancialEngine.format_money(paid_amt)}", theme_text_color='Custom', text_color=(0, 0.6, 0.2, 1), font_style='Caption', bold=True))
                    pay_row.add_widget(MDLabel(text=f'Reste : {FinancialEngine.format_money(rem)}', theme_text_color='Custom', text_color=(0.85, 0, 0, 1), font_style='Caption', bold=True, halign='right'))
                    hbox.add_widget(pay_row)
        content.add_widget(hbox)

        scroll = MDScrollView(size_hint_y=1, bar_width=dp(3))
        lst = MDList(padding=[0, dp(2)], spacing=dp(4))
        if items:
            for it in items:
                q_disp = FinancialEngine.format_qty(it['qty'])
                ib = MDBoxLayout(orientation='vertical', adaptive_height=True, padding=[dp(10), dp(8)], spacing=dp(3))
                lbl_prod = self._ui_wrap_label(self.fix_text(it['product_name']), font_style='Subtitle1', bold=True, theme_color='Primary', min_h=22)
                lbl_prod.max_lines = 2
                ib.add_widget(lbl_prod)
                if is_transfer or dt_code in AppConstants.QTY_ONLY_DOC_TYPES:
                    ib.add_widget(MDLabel(text=f'Qté : {q_disp}', theme_text_color='Custom', text_color=(0, 0.4, 0.8, 1), bold=True, adaptive_height=True))
                else:
                    lc = FinancialEngine.calculate_document_totals([it], doc_type=dt_code)
                    p_disp = FinancialEngine.format_money(it['price'])
                    ib.add_widget(MDLabel(text=f'{q_disp} x {p_disp} = {FinancialEngine.format_money(lc["final_total"])}', theme_text_color='Secondary', font_style='Body2', bold=True, adaptive_height=True))
                lst.add_widget(ib)
                lst.add_widget(MDSeparator(height=dp(1), color=(0.91, 0.92, 0.94, 1)))
        else:
            note_card = MDCard(orientation='horizontal', size_hint=(1, None), adaptive_height=True, radius=[10], elevation=0, md_bg_color=(0.96, 0.97, 0.98, 1), padding=[dp(12), dp(12)], spacing=dp(10))
            note_card.add_widget(MDIcon(icon='note-text-outline', theme_text_color='Secondary', font_size='20sp', size_hint=(None, None), size=(dp(22), dp(22)), pos_hint={'center_y': 0.5}))
            note_card.add_widget(self._ui_wrap_label(f"Note : {self.fix_text(tx.get('note', 'Opération financière'))}", font_style='Body2', bold=True, theme_color='Primary', min_h=20))
            lst.add_widget(note_card)
        scroll.add_widget(lst)
        content.add_widget(scroll)

        actions = MDBoxLayout(orientation='vertical', spacing=dp(8), size_hint=(1, None), adaptive_height=True, padding=[0, dp(4), 0, 0])
        export_row = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint=(1, None), height=dp(44))
        if dt_code not in AppConstants.INVOICE_TAX_DOC_TYPES:
            export_row.add_widget(
                self._ui_card_btn(
                    'IMP 80mm',
                    (0.08, 0.45, 0.82, 1),
                    icon='printer',
                    height=44,
                    radius=10,
                    font_size='13.5sp',
                    on_release=lambda x: threading.Thread(target=self.print_ticket_bluetooth, args=(details,), daemon=True).start(),
                )
            )
        export_row.add_widget(
            self._ui_card_btn(
                'PDF A4',
                (0.82, 0.2, 0.2, 1),
                icon='file-pdf-box',
                height=44,
                radius=10,
                font_size='13.5sp',
                on_release=lambda x: self._open_smart_pdf_dialog(op_id, is_transfer),
            )
        )
        actions.add_widget(export_row)

        if can_edit:
            actions.add_widget(
                self._ui_card_btn(
                    'MODIFIER',
                    (0.0, 0.58, 0.38, 1),
                    icon='pencil-outline',
                    height=44,
                    radius=10,
                    font_size='13.5sp',
                    on_release=lambda x: self.load_transaction_for_edit(details),
                )
            )

        actions.add_widget(
            self._ui_card_btn(
                'FERMER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                height=46,
                on_release=lambda *a: self.srv_dialog.dismiss() if getattr(self, 'srv_dialog', None) else None,
            )
        )
        content.add_widget(actions)

        self.srv_dialog = MDDialog(title='', type='custom', content_cls=content, size_hint=(0.96, None), radius=[18, 18, 18, 18], buttons=[])
        self._fit_dialog(self.srv_dialog, top_pad=14, side_pad=14, bottom_pad=12, spacer_top=0)
        self.srv_dialog.open()

    def load_transaction_for_edit(self, details: dict):
        tx = details['transaction']
        ent = details['entity']
        items = details['items']
        dt_code = tx['doc_type']

        if getattr(self, 'srv_dialog', None):
            self.srv_dialog.dismiss()

        self.editing_tx_id = int(tx['id'])
        self.editing_doc_type = dt_code
        self.current_editing_date = tx.get('date')

        if dt_code in (AppConstants.DOC_CLIENT_PAY, AppConstants.DOC_SUPPLIER_PAY):
            self.current_mode = 'client_payment' if dt_code == AppConstants.DOC_CLIENT_PAY else 'supplier_payment'
            self.selected_entity = ent
            lbl = tx.get('doc_number', '')
            self.temp_note = tx.get('note', '') if tx.get('note') != lbl else ''
            self.show_simple_payment_dialog(amount=FinancialEngine.to_decimal(tx.get('total_amount', 0)))
            return

        for d_attr in ('_active_filter_dialog', 'hist_filter_dialog', 'pending_dialog', 'entity_hist_dialog', 'mgmt_dialog', 'options_dialog'):
            dlg = getattr(self, d_attr, None)
            if dlg and dlg.parent:
                dlg.dismiss()

        doc_to_mode = {v: k for k, v in AppConstants.MODE_TO_DOC.items()}
        mode = doc_to_mode.get(dt_code, 'sale')
        self.open_mode(mode, skip_dialog=True)

        self.editing_tx_id = int(tx['id'])
        self.editing_doc_type = dt_code
        self.current_editing_date = tx.get('date')
        self.editing_payment_amount = FinancialEngine.to_decimal(tx.get('paid_amount', 0))
        self.editing_payment_method = tx.get('payment_method', AppConstants.PAY_DEFAULT)
        self.editing_global_discount = FinancialEngine.to_decimal(tx.get('global_discount', 0))
        self.selected_location = tx.get('location') or AppConstants.LOC_STORE
        self.editing_original_location = self.selected_location

        if ent:
            self.selected_entity = {
                'id': ent['id'],
                'name': ent['name'],
                'category': AppConstants.normalize_price_category(ent.get('price_category', AppConstants.PRICE_CAT_DETAIL)),
            }
            if hasattr(self, 'prod_toolbar') and mode not in ('transfer', 'initial_stock', 'manage_products'):
                ent_lbl = 'Fournisseur' if AppConstants.get_entity_type(mode) == AppConstants.ENTITY_SUPPLIER else 'Client'
                self.prod_toolbar.subtitle = f"{ent_lbl} : {self.fix_text(self.selected_entity.get('name', ''))}"
        else:
            self.selected_entity = None
            if mode == 'transfer' and hasattr(self, 'prod_toolbar'):
                src = AppConstants.LOC_LABELS[self.selected_location]
                dst = AppConstants.LOC_LABELS[AppConstants.LOC_WAREHOUSE if self.selected_location == AppConstants.LOC_STORE else AppConstants.LOC_STORE]
                self.prod_toolbar.subtitle = f'{src}  >>>  {dst}'

        self.cart = [
            {
                'id': it['product_id'],
                'name': it['product_name'],
                'qty': float(it['qty']),
                'price': float(it.get('price', 0)),
                'tva': float(it.get('tva', 0)),
                'discount_amount': float(it.get('discount_amount', 0)),
                'discount_type': it.get('discount_type', 'fixed'),
                'original_server_qty': float(it['qty']),
            }
            for it in items
        ]
        self.update_cart_button()
        self.open_cart_screen()

    def confirm_delete_transaction(self, op_id: int, is_transfer: bool = False):
        def _do_del():
            if getattr(self, 'srv_dialog', None):
                self.srv_dialog.dismiss()
            try:
                self.db.delete_transaction(op_id, is_transfer=is_transfer)
                self.notify('Opération supprimée.', 'success')
                if getattr(self, 'pending_dialog', None) and self.pending_dialog.parent:
                    self.history_page_offset = 0
                    self.history_has_more = True
                    self.is_loading_history = False
                    if hasattr(self, 'rv_history'):
                        self.rv_history.data = []
                        self.rv_history.scroll_y = 1.0
                    self.load_more_history()
                if getattr(self, 'entity_hist_dialog', None) and self.entity_hist_dialog.parent:
                    if getattr(self, 'history_target_entity', None) and self.history_target_entity.get('id'):
                        fresh_ent = self.db.get_entity_by_id(self.history_target_entity['id'])
                        if fresh_ent:
                            self.history_target_entity = fresh_ent
                    self._reload_entity_history_ui()
                self.load_more_entities(reset=True)
                self.check_and_load_stats()
            except Exception as exc:
                self.notify(str(exc), 'error')

        self.confirm_del_dialog = self._ui_confirm_dialog(
            title='Confirmer la suppression',
            message='Voulez-vous vraiment supprimer cette opération ?\nLe stock et le solde seront recalculés automatiquement.',
            confirm_text='SUPPRIMER',
            icon='trash-can-outline',
            confirm_icon='trash-can-outline',
            on_confirm=_do_del,
        )

    def _build_pdf_document(self, op_id: int, is_transfer: bool = False) -> tuple:
        details = self.db.get_operation_details(op_id, is_transfer=is_transfer)
        if not details:
            raise ValueError('Opération introuvable.')
        tx = details['transaction']
        ent = details['entity']
        items = details['items']
        dt_code = tx['doc_type']
        store_info = self.db.get_store_info()

        pdf = PDF(orientation='P', unit='mm', format='A4')
        pdf.store_info = store_info
        pdf.logo_path = self._resolve_store_logo_path(store_info)
        ent_cat = tx.get('entity_category', AppConstants.get_entity_type(dt_code))
        is_supp = ent_cat == AppConstants.ENTITY_SUPPLIER
        is_comptoir = AppConstants.is_default_entity(tx.get('entity_id'), ent_cat)
        is_financial = dt_code in (AppConstants.DOC_CLIENT_PAY, AppConstants.DOC_SUPPLIER_PAY)

        def _calc_tx_bal_impact(tx_row: dict) -> Decimal:
            d_code = tx_row['doc_type']
            t_amt = FinancialEngine.to_decimal(tx_row.get('total_amount', 0))
            p_amt = FinancialEngine.to_decimal(tx_row.get('paid_amount', 0))
            ff = AppConstants.FINANCIAL_FACTORS.get(d_code, 0)
            sf = AppConstants.STOCK_MOVEMENTS.get(d_code, 0)
            if ff == -1 and sf == 0:
                return -t_amt
            if ff == -1 and sf != 0:
                return -t_amt + p_amt
            return (t_amt * Decimal(ff)) - p_amt

        def _compute_historical_balance() -> dict | None:
            if is_comptoir or not ent or not tx.get('entity_id'):
                return None
            cur_bal = FinancialEngine.to_decimal(ent.get('balance', 0))
            conn = self.db.get_connection()
            try:
                later_rows = conn.execute(
                    '''
                    SELECT doc_type, total_amount, paid_amount
                    FROM transactions
                    WHERE entity_id = ?
                      AND entity_category = ?
                      AND (date > ? OR (date = ? AND id > ?))
                    ''',
                    (tx['entity_id'], ent_cat, tx['date'], tx['date'], tx['id']),
                ).fetchall()
            finally:
                conn.close()
            later_impact = sum((_calc_tx_bal_impact(dict(lr)) for lr in later_rows), Decimal('0.00'))
            bal_after = FinancialEngine.quantize_money(cur_bal - later_impact)
            this_impact = FinancialEngine.quantize_money(_calc_tx_bal_impact(tx))
            return {
                'old_balance': FinancialEngine.quantize_money(bal_after - this_impact),
                'transaction_amount': this_impact,
                'new_balance': bal_after,
            }

        if is_transfer:
            src = AppConstants.LOC_LABELS.get(tx.get('source_location'), 'Magasin')
            dst = AppConstants.LOC_LABELS.get(tx.get('destination_location'), 'Dépôt')
            disp_ent = f'{src} -> {dst}'
            pdf.entity_info = {'label': 'Trajet', 'name': disp_ent}
        else:
            disp_ent = ent.get('name', AppConstants.DEFAULT_SUPPLIER_NAME if is_supp else AppConstants.DEFAULT_CLIENT_NAME)
            pdf.entity_info = {
                'label': 'Fournisseur' if is_supp else 'Client',
                'name': disp_ent,
                'address': ent.get('address', ''),
                'phone': ent.get('phone', ''),
                'email': ent.get('email', ''),
                'rc': ent.get('rc', ''),
                'nif': ent.get('nif', ''),
                'nis': ent.get('nis', ''),
                'nai': ent.get('nai', ''),
                'rib': ent.get('rib', ''),
            }

        vis = AppConstants.DOC_VISUALS.get(dt_code, {'name': dt_code})
        try:
            date_str = datetime.strptime(str(tx['date'])[:19], '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y')
        except Exception:
            date_str = str(tx.get('date', ''))[:10]

        if is_financial:
            total_amt = FinancialEngine.to_decimal(tx.get('total_amount', 0))
            abs_amt = abs(total_amt)
            if dt_code == AppConstants.DOC_CLIENT_PAY:
                fin_title = 'Reçu de Versement' if total_amt >= 0 else 'Crédit Initial'
                doc_prefix = 'VER' if total_amt >= 0 else 'CI'
                recipient_sig = 'Signature Client'
            else:
                fin_title = 'Bon de Règlement' if total_amt >= 0 else 'Dette Initiale'
                doc_prefix = 'REG' if total_amt >= 0 else 'DI'
                recipient_sig = 'Signature Fournisseur'

            doc_num = f'{doc_prefix}-{op_id:05d}'
            pdf.doc_info = {
                'date': date_str,
                'doc_name_fr': fin_title.upper(),
                'doc_number': doc_num,
                'payment_method': tx.get('payment_method', ''),
                'doc_type': dt_code,
            }
            pdf.amount_in_words = f'{FinancialEngine.number_to_words_fr(abs_amt)} dinars algériens.'
            pdf.balance_data = _compute_historical_balance()
            pdf.add_page()
            pdf.draw_financial_receipt(
                fin_title,
                str(tx.get('date', ''))[:16],
                tx.get('note') or fin_title,
                abs_amt,
                recipient_sig,
            )
            doc_file_prefix = fin_title.replace(' ', '_')
        else:
            doc_num = tx.get('doc_number') or f'{dt_code}-{op_id:05d}'
            pdf.doc_info = {
                'date': date_str,
                'doc_name_fr': vis['name'].upper(),
                'doc_number': doc_num,
                'payment_method': '' if is_transfer else tx.get('payment_method', ''),
                'doc_type': dt_code,
            }

            printable_w = pdf.w - pdf.l_margin - pdf.r_margin
            table_data = []
            if dt_code in AppConstants.QTY_ONLY_DOC_TYPES:
                headers = ['N°', 'Code', 'Désignation', 'Qté']
                col_widths = [14, 32, printable_w - 76, 30]
                for i, it in enumerate(items, 1):
                    table_data.append([str(i), it.get('product_ref') or '', it['product_name'], FinancialEngine.format_qty(it['qty'])])
                pdf.totals = {}
                pdf.amount_in_words = ''
                pdf.payment_info = {}
                pdf.balance_data = None
            else:
                is_bon = dt_code in (AppConstants.DOC_BV, AppConstants.DOC_BA)
                if is_bon:
                    headers = ['N°', 'Code', 'Désignation', 'Qté', 'P.U.', 'Remise', 'Total']
                    col_widths = [10, 24, 0, 15, 25, 20, 26]
                else:
                    headers = ['N°', 'Code', 'Désignation', 'Qté', 'P.U. HT', 'Remise', 'TVA', 'Total HT']
                    col_widths = [10, 22, 0, 14, 24, 18, 14, 26]
                col_widths[2] = max(20, printable_w - sum(w for w in col_widths if w > 0))

                calc = FinancialEngine.calculate_document_totals(
                    items,
                    doc_type=dt_code,
                    global_discount=FinancialEngine.to_decimal(tx.get('global_discount', 0)),
                    payment_method=tx.get('payment_method', ''),
                    is_comptoir=is_comptoir,
                )
                for i, it in enumerate(items, 1):
                    q_s = FinancialEngine.format_qty(it['qty'])
                    p_d = FinancialEngine.to_decimal(it['price'])
                    d_amt = FinancialEngine.to_decimal(it.get('discount_amount', 0))
                    u_disc = d_amt if it.get('discount_type') == 'fixed' else p_d * (d_amt / Decimal('100'))
                    line_ht = FinancialEngine.quantize_money((p_d - u_disc) * FinancialEngine.to_decimal(it['qty']))
                    row = [str(i), it.get('product_ref') or '', it['product_name'], q_s, FinancialEngine.format_money(p_d, False), FinancialEngine.format_money(d_amt, False)]
                    if not is_bon:
                        row.append(f"{float(it.get('tva', 0)):g}%")
                    row.append(FinancialEngine.format_money(line_ht, False))
                    table_data.append(row)

                pdf.totals = {
                    'total_ht': calc['total_ht_final'],
                    'total_tva': calc['total_tva_final'],
                    'total_discount': calc['total_all_discounts_final'],
                    'stamp_duty': calc['stamp_duty'] if calc['stamp_duty'] > 0 else None,
                    'final_total': calc['final_total'],
                }
                pdf.amount_in_words = f"{FinancialEngine.number_to_words_fr(calc['final_total'])} dinars algériens."
                pdf.payment_info = {'amount': tx.get('paid_amount', 0)}

                pdf.balance_data = None
                if self.db.get_bool_setting('show_balance_in_pdf', False) and dt_code in (AppConstants.DOC_BV, AppConstants.DOC_BA) and not is_comptoir and ent:
                    pdf.balance_data = _compute_historical_balance()

            pdf.add_page()
            pdf.draw_table_with_fill(headers, table_data, col_widths)
            doc_file_prefix = vis['name'].replace(' ', '_')

        safe_client = ''.join(c for c in disp_ent if c.isalnum() or c in (' ', '-', '_')).strip().replace(' ', '_') or 'Client'
        safe_num = str(doc_num).replace('/', '-').replace('\\', '-').replace(':', '')
        fname = f'{doc_file_prefix}_{safe_num}_{safe_client}.pdf'

        temp_pdf = os.path.join(self.user_data_dir, f'temp_build_{int(time.time())}_{random.randint(100, 999)}.pdf')
        try:
            pdf.output(temp_pdf)
            with open(temp_pdf, 'rb') as f:
                pdf_bytes = f.read()
        finally:
            if os.path.exists(temp_pdf):
                try:
                    os.remove(temp_pdf)
                except Exception:
                    pass
        return fname, pdf_bytes

    def _compute_pdf_signature(self, details: dict) -> str:
        store_info = self.db.get_store_info()
        show_bal = self.db.get_bool_setting('show_balance_in_pdf', False)
        payload = {
            'tx': details.get('transaction', {}),
            'ent': details.get('entity', {}),
            'items': details.get('items', []),
            'store': store_info,
            'show_bal': show_bal,
        }
        raw_json = json.dumps(payload, sort_keys=True, default=str)
        return hashlib.sha256(raw_json.encode('utf-8')).hexdigest()

    def _check_pdf_exists(self, saved_path: str, uri_str: str) -> bool:
        if uri_str and platform == 'android' and PythonActivity is not None:
            try:
                from jnius import autoclass
                uri_obj = autoclass('android.net.Uri').parse(uri_str)
                resolver = PythonActivity.mActivity.getContentResolver()
                stream = resolver.openInputStream(uri_obj)
                if stream is not None:
                    stream.close()
                    return True
            except Exception:
                return False
        return bool(saved_path and os.path.exists(saved_path))

    def _get_temp_share_path(self, fname: str) -> str:
        base_dir = self.user_data_dir
        if platform == 'android' and PythonActivity is not None:
            try:
                ext_cache = PythonActivity.mActivity.getExternalCacheDir()
                if ext_cache is not None:
                    base_dir = ext_cache.getAbsolutePath()
            except Exception:
                pass
        share_dir = os.path.join(base_dir, 'share_tmp')
        os.makedirs(share_dir, exist_ok=True)
        for old_f in os.listdir(share_dir):
            try:
                os.remove(os.path.join(share_dir, old_f))
            except Exception:
                pass
        return os.path.join(share_dir, fname)

    def generate_pdf_report(self, op_id: int, is_transfer: bool = False):
        self._open_smart_pdf_dialog(op_id, is_transfer=is_transfer)

    def _open_smart_pdf_dialog(self, op_id: int, is_transfer: bool = False):
        if getattr(self, 'smart_pdf_dialog', None):
            try:
                self.smart_pdf_dialog.dismiss()
            except Exception:
                pass

        details = self.db.get_operation_details(op_id, is_transfer=is_transfer)
        if not details:
            self.notify('Opération introuvable.', 'error')
            return

        doc_sig = self._compute_pdf_signature(details)
        state_key = f"pdf_state_{'TR' if is_transfer else 'TX'}_{op_id}"
        cached = {}
        raw_state = self.db.get_setting(state_key, '')
        if raw_state:
            try:
                cached = json.loads(raw_state)
            except Exception:
                cached = {}

        is_downloaded_fresh = (
            isinstance(cached, dict)
            and cached.get('sig') == doc_sig
            and self._check_pdf_exists(cached.get('path', ''), cached.get('uri', ''))
        )
        tx = details['transaction']
        doc_label = tx.get('doc_number') or f"OP-{op_id:05d}"

        state_box = {
            'ready': is_downloaded_fresh,
            'path': cached.get('path', '') if is_downloaded_fresh else '',
            'uri_str': cached.get('uri', '') if is_downloaded_fresh else '',
            'fname': cached.get('fname', '') if is_downloaded_fresh else '',
            'busy': False,
        }

        content = MDBoxLayout(orientation='vertical', spacing=dp(12), size_hint_y=None, adaptive_height=True, padding=[0, dp(2), 0, dp(2)])

        badge = self._ui_icon_pill(
            'file-check-outline' if is_downloaded_fresh else 'file-pdf-box',
            (0.90, 0.98, 0.92, 1) if is_downloaded_fresh else (1, 0.93, 0.93, 1),
            (0.04, 0.62, 0.25, 1) if is_downloaded_fresh else (0.82, 0.16, 0.16, 1),
            size_dp=56,
            radius=28,
            icon_size='28sp',
        )
        badge.pos_hint = {'center_x': 0.5}
        content.add_widget(badge)

        content.add_widget(MDLabel(text='Document PDF A4', halign='center', font_style='H6', bold=True, theme_text_color='Primary', size_hint_y=None, height=dp(28)))

        info_card = MDCard(orientation='vertical', padding=[dp(12), dp(10)], spacing=dp(6), radius=[12], elevation=0, md_bg_color=(0.95, 0.97, 1, 1), line_color=(0.84, 0.88, 0.95, 1), line_width=1, size_hint_y=None, adaptive_height=True)
        lbl_status = self._ui_wrap_label(
            f"Déjà téléchargé : {self.fix_text(state_box['fname'])}" if is_downloaded_fresh else f'Document {self.fix_text(doc_label)} prêt à télécharger ou partager.',
            font_style='Caption',
            bold=True,
            theme_color='Secondary',
            min_h=18,
        )
        info_card.add_widget(lbl_status)
        content.add_widget(info_card)

        btn_primary = self._ui_card_btn(
            'OUVRIR' if is_downloaded_fresh else 'TÉLÉCHARGER',
            (0.09, 0.44, 0.84, 1) if is_downloaded_fresh else (0.04, 0.62, 0.25, 1),
            icon='file-eye-outline' if is_downloaded_fresh else 'download',
            size_hint_x=0.54,
            height=48,
            font_size='12.5sp',
            elevation=1,
        )
        btn_share = self._ui_card_btn(
            'PARTAGER',
            (0.0, 0.55, 0.62, 1),
            icon='share-variant',
            size_hint_x=0.46,
            height=48,
            font_size='12.5sp',
            elevation=1,
        )

        def _parse_uri(uri_str: str):
            if uri_str and platform == 'android':
                try:
                    from jnius import autoclass
                    return autoclass('android.net.Uri').parse(uri_str)
                except Exception:
                    return None
            return None

        def _mark_downloaded(fname: str, saved_path: str, saved_uri):
            uri_str = str(saved_uri.toString()) if saved_uri is not None else ''
            self.db.set_setting(
                state_key,
                json.dumps({'sig': doc_sig, 'path': saved_path, 'uri': uri_str, 'fname': fname}),
            )
            state_box['ready'] = True
            state_box['path'] = saved_path
            state_box['uri_str'] = uri_str
            state_box['fname'] = fname

            btn_primary.md_bg_color = (0.09, 0.44, 0.84, 1)
            btn_primary._lbl.text = 'OUVRIR'
            btn_primary._ico.icon = 'file-eye-outline'
            if hasattr(badge, '_bg_color_instr') and badge._bg_color_instr:
                badge._bg_color_instr.rgba = (0.90, 0.98, 0.92, 1)
            else:
                badge.md_bg_color = (0.90, 0.98, 0.92, 1)
            badge._ico.icon = 'file-check-outline'
            badge._ico.text_color = (0.04, 0.62, 0.25, 1)
            lbl_status.text = f'Enregistré dans Téléchargements : {self.fix_text(fname)}'

        def _on_primary_tap(*args):
            if state_box['busy']:
                return
            if state_box['ready']:
                self._share_or_open_exported_file(state_box['path'], _parse_uri(state_box['uri_str']), 'application/pdf', action='view')
                return

            state_box['busy'] = True
            try:
                fname, pdf_bytes = self._build_pdf_document(op_id, is_transfer=is_transfer)
                saved_path, saved_uri = self._save_bytes_to_downloads(fname, pdf_bytes, 'application/pdf')
                _mark_downloaded(fname, saved_path, saved_uri)
                self.notify(f'PDF téléchargé : {fname}', 'success')
            except Exception as exc:
                self.notify(f'Erreur PDF : {exc}', 'error')
            finally:
                state_box['busy'] = False

        def _on_share_tap(*args):
            if state_box['busy']:
                return
            state_box['busy'] = True
            try:
                if state_box['ready'] and self._check_pdf_exists(state_box['path'], state_box['uri_str']):
                    self._share_or_open_exported_file(state_box['path'], _parse_uri(state_box['uri_str']), 'application/pdf', action='share')
                    return

                fname, pdf_bytes = self._build_pdf_document(op_id, is_transfer=is_transfer)
                if platform == 'android':
                    saved_path, saved_uri = self._save_bytes_to_downloads(fname, pdf_bytes, 'application/pdf')
                    _mark_downloaded(fname, saved_path, saved_uri)
                    self._share_or_open_exported_file(saved_path, saved_uri, 'application/pdf', action='share')
                else:
                    tmp_path = self._get_temp_share_path(fname)
                    with open(tmp_path, 'wb') as f:
                        f.write(pdf_bytes)
                    self._share_or_open_exported_file(tmp_path, None, 'application/pdf', action='share')
            except Exception as exc:
                self.notify(f'Erreur partage PDF : {exc}', 'error')
            finally:
                state_box['busy'] = False

        btn_primary.bind(on_release=_on_primary_tap)
        btn_share.bind(on_release=_on_share_tap)

        actions_row = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint=(1, None), height=dp(48))
        actions_row.add_widget(btn_primary)
        actions_row.add_widget(btn_share)
        content.add_widget(actions_row)

        content.add_widget(
            self._ui_card_btn(
                'FERMER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                height=44,
                on_release=lambda x: self.smart_pdf_dialog.dismiss() if getattr(self, 'smart_pdf_dialog', None) else None,
            )
        )

        self.smart_pdf_dialog = MDDialog(
            title='',
            type='custom',
            content_cls=content,
            radius=[20, 20, 20, 20],
            size_hint=(0.94, None),
            buttons=[],
        )
        self._fit_dialog(self.smart_pdf_dialog, top_pad=16, side_pad=14, bottom_pad=14, spacer_top=0)
        self.smart_pdf_dialog.open()

    # --------------------------------------------------------------------------
    # 7.11 IMPRESSION BLUETOOTH ESC/POS 80 MM (AVEC LOGO MAGASIN)
    # --------------------------------------------------------------------------
    def open_bluetooth_selector(self, instance):
        if platform != 'android' or BluetoothAdapter is None:
            self.notify('Disponible uniquement sur Android.', 'error')
            return
        try:
            adapter = BluetoothAdapter.getDefaultAdapter()
            if not adapter or not adapter.isEnabled():
                self.notify('Veuillez activer le Bluetooth.', 'error')
                return
            paired = adapter.getBondedDevices().toArray()
            current_mac = str(self.db.get_setting('printer_mac', '') or '').strip()

            content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, height=min(dp(420), Window.height * 0.68), padding=[0, dp(4), 0, 0])
            scroll = MDScrollView(size_hint_y=1, bar_width=dp(3))
            lst = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(8), padding=[0, dp(2), dp(2), dp(4)])

            if not paired:
                empty_card = MDCard(orientation='vertical', size_hint=(1, None), height=dp(150), radius=[14], elevation=0, md_bg_color=(0.97, 0.98, 0.99, 1), line_color=(0.88, 0.90, 0.94, 1), line_width=1, padding=dp(16), spacing=dp(8))
                empty_card.add_widget(MDBoxLayout())
                empty_card.add_widget(MDIcon(icon='bluetooth-off', theme_text_color='Hint', font_size='38sp', halign='center', pos_hint={'center_x': 0.5}))
                empty_card.add_widget(MDLabel(text='Aucun appareil associé (Paired)', halign='center', font_style='Subtitle2', bold=True, theme_text_color='Secondary', adaptive_height=True))
                empty_card.add_widget(MDBoxLayout())
                lst.add_widget(empty_card)
            else:
                for dev in paired:
                    dn = str(dev.getName() or 'Imprimante Bluetooth')
                    dm = str(dev.getAddress() or '')
                    is_sel = bool(current_mac and dm == current_mac)
                    card = MDCard(
                        orientation='horizontal',
                        size_hint=(1, None),
                        height=dp(62),
                        radius=[12],
                        elevation=0,
                        md_bg_color=(0.90, 0.95, 1, 1) if is_sel else (0.97, 0.98, 0.99, 1),
                        line_color=(0.09, 0.44, 0.84, 0.65) if is_sel else (0.88, 0.90, 0.94, 1),
                        line_width=1.3 if is_sel else 1,
                        padding=[dp(12), dp(8)],
                        spacing=dp(10),
                        ripple_behavior=True,
                        on_release=lambda x, n=dn, m=dm: self.select_printer(n, m),
                    )
                    card.add_widget(
                        self._ui_icon_pill(
                            'printer-wireless',
                            (0.09, 0.44, 0.84, 0.14) if is_sel else (0.90, 0.92, 0.96, 1),
                            (0.09, 0.44, 0.84, 1) if is_sel else (0.35, 0.40, 0.48, 1),
                            size_dp=40,
                            icon_size='21sp',
                        )
                    )
                    info = MDBoxLayout(orientation='vertical', size_hint_x=1, adaptive_height=True, spacing=dp(2), pos_hint={'center_y': 0.5})
                    info.add_widget(MDLabel(text=dn, font_style='Subtitle2', bold=True, theme_text_color='Primary', shorten=True, adaptive_height=True))
                    info.add_widget(MDLabel(text=dm, font_style='Caption', theme_text_color='Secondary', shorten=True, adaptive_height=True))
                    card.add_widget(info)
                    card.add_widget(
                        MDIcon(
                            icon='check-circle' if is_sel else 'chevron-right',
                            theme_text_color='Custom' if is_sel else 'Hint',
                            text_color=(0.09, 0.44, 0.84, 1) if is_sel else (0.5, 0.5, 0.5, 1),
                            font_size='21sp',
                            size_hint_x=None,
                            width=dp(22),
                            pos_hint={'center_y': 0.5},
                        )
                    )
                    lst.add_widget(card)

            scroll.add_widget(lst)
            content.add_widget(scroll)
            content.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), height=46, on_release=lambda x: self.bt_dialog.dismiss()))

            self.bt_dialog = MDDialog(title='Choisir Imprimante', type='custom', content_cls=content, size_hint=(0.92, None), radius=[18, 18, 18, 18], buttons=[])
            self._fit_dialog(self.bt_dialog)
            self.bt_dialog.open()
        except Exception as exc:
            self.notify(f'Erreur Bluetooth : {exc}', 'error')

    def select_printer(self, name: str, mac: str):
        self.db.set_setting('printer_mac', mac)
        self.db.set_setting('printer_name', name)
        if getattr(self, 'bt_dialog', None):
            self.bt_dialog.dismiss()
        self.notify(f'Imprimante sélectionnée : {name}', 'success')
        if getattr(self, 'settings_menu_dialog', None):
            self.settings_menu_dialog.dismiss()
            self.open_settings_menu()

    def clear_printer_selection(self, instance):
        self.db.set_setting('printer_mac', '')
        self.db.set_setting('printer_name', '')
        self.notify('Imprimante dissociée.', 'info')
        if getattr(self, 'settings_menu_dialog', None):
            self.settings_menu_dialog.dismiss()
            self.open_settings_menu()

    @staticmethod
    def get_wrapped_text(text: str, font, max_width: int) -> list:
        lines = []
        if not text:
            return lines
        cur = []
        for word in text.split(' '):
            cur.append(word)
            bb = font.getbbox(' '.join(cur))
            if (bb[2] - bb[0]) > max_width:
                if len(cur) == 1:
                    lines.append(cur[0])
                    cur = []
                else:
                    cur.pop()
                    lines.append(' '.join(cur))
                    cur = [word]
        if cur:
            lines.append(' '.join(cur))
        return lines

    @staticmethod
    def get_image_raster_data(image: Image.Image) -> list:
        max_w = 576
        if image.width > max_w:
            image = image.resize((max_w, int(image.height * (max_w / float(image.width)))), Image.Resampling.LANCZOS)
        if image.width % 8 != 0:
            image = image.crop((0, 0, (image.width // 8) * 8, image.height))
        image = image.convert('1')
        w, h = image.size
        bpr = w // 8
        band_h = 120
        chunks = []
        for y in range(0, h, band_h):
            ch = min(band_h, h - y)
            band = image.crop((0, y, w, y + ch))
            cmd = b'\x1dv0\x00' + bytes([bpr % 256, bpr // 256, ch % 256, ch // 256])
            chunks.append(cmd + bytearray(b ^ 255 for b in band.tobytes()))
        return chunks

    def create_receipt_image(self, details: dict) -> Image.Image:
        tx = details['transaction']
        ent = details['entity']
        items = details['items']
        dt_code = tx['doc_type']
        is_transfer = dt_code == AppConstants.DOC_TR
        is_simple = dt_code in (AppConstants.DOC_CLIENT_PAY, AppConstants.DOC_SUPPLIER_PAY)

        paper_w = 576
        margin = 10
        img = Image.new('RGB', (paper_w, 4500), (255, 255, 255))
        draw = ImageDraw.Draw(img)
        try:
            f_reg = ImageFont.truetype(FONT_FILE, 22)
            f_med = ImageFont.truetype(FONT_FILE, 26)
            f_large = ImageFont.truetype(FONT_FILE, 38)
        except Exception:
            f_reg = f_med = f_large = ImageFont.load_default()

        def _line(text, y_pos, font_o, align='left'):
            if not text:
                return y_pos
            vis = ar_visual(str(text))
            bb = font_o.getbbox(vis)
            tw, th = bb[2] - bb[0], bb[3] - bb[1]
            xp = margin if align == 'left' else ((paper_w - tw) // 2 if align == 'center' else paper_w - tw - margin)
            draw.text((xp, y_pos), vis, font=font_o, fill=(0, 0, 0))
            return y_pos + th + 8

        def _sep(y_pos):
            draw.line([(margin, y_pos), (paper_w - margin, y_pos)], fill=(0, 0, 0), width=2)
            return y_pos + 10

        def _lr(left, right, font_o, y_pos, bold=False):
            l, r = ar_visual(str(left)), ar_visual(str(right))
            br, bl = font_o.getbbox(r), font_o.getbbox(l)
            xr = paper_w - (br[2] - br[0]) - margin
            draw.text((margin, y_pos), l, font=font_o, fill=(0, 0, 0))
            draw.text((xr, y_pos), r, font=font_o, fill=(0, 0, 0))
            if bold:
                draw.text((margin + 1, y_pos), l, font=font_o, fill=(0, 0, 0))
                draw.text((xr + 1, y_pos), r, font=font_o, fill=(0, 0, 0))
            return max(br[3] - br[1], bl[3] - bl[1], 30) + 8

        store_info = self.db.get_store_info()
        y = 10
        logo_path = self._resolve_store_logo_path(store_info)
        if logo_path:
            try:
                with Image.open(logo_path) as l_im:
                    if l_im.mode in ('RGBA', 'P', 'LA'):
                        bg = Image.new('RGB', l_im.size, (255, 255, 255))
                        rgba = l_im.convert('RGBA')
                        bg.paste(rgba, mask=rgba.split()[3])
                        l_im = bg
                    else:
                        l_im = l_im.convert('RGB')
                    l_im.thumbnail((220, 140), Image.Resampling.LANCZOS)
                    lw, lh = l_im.size
                    img.paste(l_im, ((paper_w - lw) // 2, y))
                    y += lh + 10
            except Exception:
                pass

        y = _line(store_info.get('store_name', AppConstants.APP_NAME), y, f_large, 'center')
        if store_info.get('store_address'):
            y = _line(store_info['store_address'], y, f_reg, 'center')
        if store_info.get('store_phone'):
            y = _line(f"Tél : {store_info['store_phone']}", y, f_reg, 'center')
        y = _sep(y + 5)

        vis = AppConstants.DOC_VISUALS.get(dt_code, {'name': dt_code})
        title = str(tx.get('doc_number', vis['name'])).upper() if is_simple else vis['name'].upper()
        y = _line(title, y, f_large, 'center')
        y = _sep(y + 5)

        if not is_simple and tx.get('doc_number'):
            y = _line(f"N° : {tx['doc_number']}", y, f_med, 'left')
        y = _line(f"Date : {str(tx.get('date', ''))[:16]}", y, f_reg, 'left')

        ent_cat = tx.get('entity_category', AppConstants.get_entity_type(dt_code))
        is_comptoir = AppConstants.is_default_entity(tx.get('entity_id'), ent_cat)
        if is_transfer:
            src = AppConstants.LOC_LABELS.get(tx.get('source_location'), 'Magasin')
            dst = AppConstants.LOC_LABELS.get(tx.get('destination_location'), 'Dépôt')
            y = _line(f'Trajet : {src} -> {dst}', y, f_med, 'left')
        else:
            lbl_e = 'Fournisseur' if ent_cat == AppConstants.ENTITY_SUPPLIER else 'Client'
            e_name = ent.get('name', AppConstants.DEFAULT_CLIENT_NAME)
            y = _line(f'{lbl_e} : {e_name}', y, f_med, 'left')
        y = _sep(y + 5)

        if is_simple:
            amt = abs(FinancialEngine.to_decimal(tx.get('total_amount', 0)))
            y = _line('MONTANT :', y + 20, f_large, 'center')
            y = _line(FinancialEngine.format_money(amt), y, f_large, 'center')
            note = str(tx.get('note', '')).strip()
            if note and note.upper() != title:
                y = _line(f'Note : {note}', y + 10, f_reg, 'center')
            y += 20
            draw.line([(margin + 100, y), (paper_w - margin - 100, y)], fill=(0, 0, 0), width=1)
            y = _line('Merci de votre fidélité !', y + 15, f_med, 'center')
            return img.crop((0, 0, paper_w, y + 120))

        for it in items:
            for ln in self.get_wrapped_text(it['product_name'], f_med, paper_w - 2 * margin):
                y = _line(ln, y, f_med, 'left')
            q_s = FinancialEngine.format_qty(it['qty'])
            if dt_code in AppConstants.QTY_ONLY_DOC_TYPES:
                y = _line(f'Qté : {q_s}', y, f_large, 'center')
            else:
                lc = FinancialEngine.calculate_document_totals([it], doc_type=dt_code)
                p_s = FinancialEngine.format_money(it['price'], False)
                t_rate = FinancialEngine.to_decimal(it.get('tva', 0))
                left_s = f'{q_s} x {p_s}' + (f' (TVA {int(t_rate)}%)' if (t_rate > 0 and dt_code in AppConstants.INVOICE_TAX_DOC_TYPES) else '')
                y += _lr(left_s, FinancialEngine.format_money(lc['final_total'], False), f_reg, y)
            draw.line([(margin + 50, y - 2), (paper_w - margin - 50, y - 2)], fill=(210, 210, 210), width=1)
            y += 6

        y = _sep(y + 10)
        if dt_code not in AppConstants.QTY_ONLY_DOC_TYPES:
            calc = FinancialEngine.calculate_document_totals(
                items,
                doc_type=dt_code,
                global_discount=FinancialEngine.to_decimal(tx.get('global_discount', 0)),
                payment_method=tx.get('payment_method', ''),
                is_comptoir=is_comptoir,
            )
            if calc['total_tva_final'] > 0:
                y += _lr('Total HT :', FinancialEngine.format_money(calc['total_ht_final']), f_med, y)
                y += _lr('Total TVA :', FinancialEngine.format_money(calc['total_tva_final']), f_med, y)
            if calc['stamp_duty'] > 0:
                y += _lr('Droit Timbre :', FinancialEngine.format_money(calc['stamp_duty']), f_med, y)
            y += _lr('TOTAL :', FinancialEngine.format_money(calc['final_total']), f_large, y + 6, True)
            if not is_comptoir and AppConstants.FINANCIAL_FACTORS.get(dt_code, 0) == 1:
                paid_d = FinancialEngine.to_decimal(tx.get('paid_amount', 0))
                rem = FinancialEngine.quantize_money(calc['final_total'] - paid_d)
                y += _lr('VERSEMENT :', FinancialEngine.format_money(paid_d), f_med, y + 4)
                y += _lr('RESTE :', FinancialEngine.format_money(max(Decimal('0'), rem)), f_med, y, True)

        y += 25
        draw.line([(margin + 100, y), (paper_w - margin - 100, y)], fill=(0, 0, 0), width=1)
        y = _line('Merci de votre fidélité !', y + 15, f_med, 'center')
        return img.crop((0, 0, paper_w, y + 120))

    def print_ticket_bluetooth(self, details: dict):
        if platform != 'android' or BluetoothAdapter is None:
            return
        mac = str(self.db.get_setting('printer_mac', '') or '').strip()
        if not mac:
            self.notify('Imprimante Bluetooth non configurée.', 'error')
            return
        sock = None
        try:
            adapter = BluetoothAdapter.getDefaultAdapter()
            if not adapter or not adapter.isEnabled():
                self.notify('Bluetooth désactivé.', 'error')
                return
            dev = adapter.getRemoteDevice(mac)
            uuid = UUID.fromString('00001101-0000-1000-8000-00805F9B34FB')
            chunks = self.get_image_raster_data(self.create_receipt_image(details))
            sock = dev.createRfcommSocketToServiceRecord(uuid)
            sock.connect()
            time.sleep(0.3)
            out = sock.getOutputStream()
            out.write(b'\x1b@')
            out.flush()
            time.sleep(0.1)
            for c in chunks:
                out.write(c)
                out.flush()
                time.sleep(0.15)
            out.write(b'\n\n\n\n\n')
            out.flush()
            time.sleep(0.2)
            out.write(b'\x1dV\x00')
            out.flush()
            time.sleep(0.5)
            sock.close()
        except Exception:
            try:
                if sock:
                    sock.close()
            except Exception:
                pass
            self.notify("Erreur d'impression Bluetooth.", 'error')

    # --------------------------------------------------------------------------
    # 7.12 SCANNER CODE-BARRES CAMÉRA (SIMPLE & CONTINU)
    # --------------------------------------------------------------------------
    def open_barcode_scanner(self, instance):
        self.temp_scanned_cart = []
        self.potential_code = None
        self.consecutive_frames = 0
        self._consensus_ts = 0.0
        self.last_scan_time = 0.0
        self._last_any_scan_time = 0.0
        self._recent_scan_times = {}
        self._frame_scan_tick = 0
        self.is_detecting_frame = False
        self.is_scanner_paused = False
        self.scanner_sound_enabled = getattr(self, 'scanner_sound_enabled', True)
        self._current_zoom = 1.0
        self._zoom_initialized = False
        self._zoom_ratios = []
        self._torch_active = False
        self._torch_supported = False
        self._scanner_session_id = getattr(self, '_scanner_session_id', 0) + 1
        if instance is not None:
            self.target_scan_field = None
        if not zbar_decode:
            self.notify('Librairie pyzbar manquante.', 'error')
            return
        if platform == 'android':
            from android.permissions import Permission, request_permissions
            request_permissions(
                [Permission.CAMERA],
                lambda p, g: Clock.schedule_once(lambda dt: self._launch_camera_widget(), 0.1) if (g and g[0]) else self.notify('Permission caméra refusée.', 'error'),
            )
        else:
            self._launch_camera_widget()

    def _get_retail_zbar_symbols(self):
        if hasattr(self, '_cached_zbar_symbols'):
            return self._cached_zbar_symbols
        try:
            from pyzbar.pyzbar import ZBarSymbol
            syms = []
            for name in ('EAN13', 'EAN8', 'UPCA', 'UPCE', 'CODE128', 'CODE39', 'CODE93', 'I25', 'CODABAR', 'QRCODE'):
                sym = getattr(ZBarSymbol, name, None)
                if sym is not None:
                    syms.append(sym)
            self._cached_zbar_symbols = syms if syms else None
        except Exception:
            self._cached_zbar_symbols = None
        return self._cached_zbar_symbols

    def _get_android_camera(self):
        if platform != 'android' or not getattr(self, 'camera_widget', None):
            return None
        try:
            cam_core = getattr(self.camera_widget, '_camera', None)
            if cam_core is None:
                return None
            return getattr(cam_core, '_android_camera', None) or getattr(cam_core, '_camera', None)
        except Exception:
            return None

    def _optimize_camera_params(self):
        if not getattr(self, 'scan_dialog', None) or not self.scan_dialog.parent:
            return
        acam = self._get_android_camera()
        if acam is None:
            return
        try:
            params = acam.getParameters()
            if params is None:
                return

            params_changed = False
            try:
                f_modes = params.getSupportedFocusModes()
                if f_modes is not None:
                    modes_list = [str(m) for m in f_modes.toArray()]
                    cur_focus = str(params.getFocusMode() or '')
                    if 'continuous-picture' in modes_list and cur_focus != 'continuous-picture':
                        params.setFocusMode('continuous-picture')
                        params_changed = True
                    elif 'continuous-video' in modes_list and cur_focus not in ('continuous-picture', 'continuous-video'):
                        params.setFocusMode('continuous-video')
                        params_changed = True
                    elif 'auto' in modes_list and not cur_focus:
                        params.setFocusMode('auto')
                        params_changed = True
            except Exception:
                pass

            try:
                if params.isZoomSupported():
                    ratios_raw = params.getZoomRatios()
                    if ratios_raw is not None:
                        self._zoom_ratios = [int(r) for r in ratios_raw.toArray()]
                    else:
                        self._zoom_ratios = []
                    if not self._zoom_initialized:
                        self._zoom_initialized = True
                        self._current_zoom = 1.0
                    if getattr(self, 'btn_scan_zoom', None):
                        if self._zoom_ratios and max(self._zoom_ratios) >= 150:
                            self.btn_scan_zoom.opacity = 1
                            self.btn_scan_zoom.disabled = False
                            z_val = self._current_zoom
                            self.btn_scan_zoom.text = f'{int(z_val)}x' if float(z_val).is_integer() else f'{z_val:.1f}x'
                        else:
                            self.btn_scan_zoom.opacity = 0
                            self.btn_scan_zoom.disabled = True
                elif getattr(self, 'btn_scan_zoom', None):
                    self.btn_scan_zoom.opacity = 0
                    self.btn_scan_zoom.disabled = True
            except Exception:
                pass

            try:
                fl_modes = params.getSupportedFlashModes()
                fl_list = [str(m) for m in fl_modes.toArray()] if fl_modes is not None else []
                self._torch_supported = 'torch' in fl_list
                if getattr(self, 'btn_scan_torch', None):
                    self.btn_scan_torch.opacity = 1 if self._torch_supported else 0
                    self.btn_scan_torch.disabled = not self._torch_supported
            except Exception:
                pass

            if params_changed:
                acam.setParameters(params)
        except Exception:
            pass

    def _trigger_tap_to_focus(self, touch_pos):
        if getattr(self, 'is_scanner_paused', False):
            return
        if getattr(self, '_focus_ring_line', None) and getattr(self, '_focus_ring_color', None):
            self._focus_ring_line.circle = (touch_pos[0], touch_pos[1], dp(26))
            self._focus_ring_color.a = 1.0
            if getattr(self, '_focus_ring_ev', None):
                self._focus_ring_ev.cancel()
            self._focus_ring_ev = Clock.schedule_once(
                lambda dt: setattr(self._focus_ring_color, 'a', 0.0) if getattr(self, '_focus_ring_color', None) else None,
                0.55,
            )

        acam = self._get_android_camera()
        if acam is None:
            return
        try:
            params = acam.getParameters()
            if params is None:
                return
            f_modes = params.getSupportedFocusModes()
            modes_list = [str(m) for m in f_modes.toArray()] if f_modes is not None else []
            if 'auto' in modes_list:
                params.setFocusMode('auto')
                acam.setParameters(params)
            acam.autoFocus(None)

            def _restore_cont(dt):
                c_obj = self._get_android_camera()
                if c_obj is None:
                    return
                try:
                    p2 = c_obj.getParameters()
                    fm2 = p2.getSupportedFocusModes()
                    ml2 = [str(m) for m in fm2.toArray()] if fm2 is not None else []
                    if 'continuous-picture' in ml2:
                        p2.setFocusMode('continuous-picture')
                        c_obj.setParameters(p2)
                except Exception:
                    pass

            Clock.schedule_once(_restore_cont, 0.65)
        except Exception:
            pass

    def _toggle_camera_zoom(self, instance=None):
        acam = self._get_android_camera()
        if acam is None or not getattr(self, '_zoom_ratios', None):
            return
        try:
            params = acam.getParameters()
            if params is None or not params.isZoomSupported():
                return
            max_ratio = max(self._zoom_ratios)
            if self._current_zoom < 1.4:
                target = min(1.5, max_ratio / 100.0)
            elif self._current_zoom < 1.9 and max_ratio >= 200:
                target = 2.0
            elif self._current_zoom < 2.8 and max_ratio >= 300:
                target = 3.0
            else:
                target = 1.0

            target_hundred = int(round(target * 100))
            best_idx = 0
            best_diff = 999999
            for idx, r_val in enumerate(self._zoom_ratios):
                d = abs(r_val - target_hundred)
                if d < best_diff:
                    best_diff = d
                    best_idx = idx

            params.setZoom(best_idx)
            acam.setParameters(params)
            self._current_zoom = target
            if getattr(self, 'btn_scan_zoom', None):
                self.btn_scan_zoom.text = f'{int(target)}x' if float(target).is_integer() else f'{target:.1f}x'
                self.btn_scan_zoom.md_bg_color = (0.15, 0.45, 0.85, 0.85) if target > 1.0 else (0.15, 0.15, 0.18, 0.75)
        except Exception:
            pass

    def _toggle_camera_torch(self, instance=None):
        if not getattr(self, '_torch_supported', False):
            self.notify('Flash non disponible', 'warning')
            return
        acam = self._get_android_camera()
        if acam is None:
            return
        try:
            params = acam.getParameters()
            if params is None:
                return
            self._torch_active = not getattr(self, '_torch_active', False)
            params.setFlashMode('torch' if self._torch_active else 'off')
            acam.setParameters(params)
            if getattr(self, 'btn_scan_torch', None):
                self.btn_scan_torch.icon = 'flashlight' if self._torch_active else 'flashlight-off'
                self.btn_scan_torch.md_bg_color = (0.95, 0.65, 0.05, 0.9) if self._torch_active else (0.15, 0.15, 0.18, 0.75)
        except Exception:
            self._torch_active = False
            self.notify('Flash non disponible', 'warning')

    def _toggle_scanner_sound(self, instance=None):
        self.scanner_sound_enabled = not getattr(self, 'scanner_sound_enabled', True)
        if getattr(self, 'btn_scan_sound', None):
            self.btn_scan_sound.icon = 'volume-high' if self.scanner_sound_enabled else 'volume-off'
            self.btn_scan_sound.md_bg_color = (0.15, 0.45, 0.85, 0.85) if self.scanner_sound_enabled else (0.15, 0.15, 0.18, 0.75)

    def _toggle_camera_pause(self, instance=None):
        if not getattr(self, 'camera_widget', None):
            return
        self.is_scanner_paused = not getattr(self, 'is_scanner_paused', False)
        if self.is_scanner_paused:
            if getattr(self, '_torch_active', False):
                self._toggle_camera_torch()
            self.camera_widget.play = False
            self.potential_code = None
            self.consecutive_frames = 0
            if getattr(self, 'btn_scan_pause', None):
                self.btn_scan_pause.icon = 'play'
                self.btn_scan_pause.md_bg_color = (0.05, 0.65, 0.35, 0.9)
            if getattr(self, '_laser_color', None):
                self._laser_color.a = 0.0
        else:
            self.scanner_start_time = time.time()
            self.camera_widget.play = True
            self._zoom_initialized = False
            if getattr(self, 'btn_scan_pause', None):
                self.btn_scan_pause.icon = 'pause'
                self.btn_scan_pause.md_bg_color = (0.15, 0.15, 0.18, 0.75)
            if getattr(self, '_laser_color', None):
                self._laser_color.a = 0.95
            Clock.schedule_once(lambda dt: self._optimize_camera_params(), 0.4)

    def _update_camera_preview_layout(self, *args):
        preview_clip = getattr(self, 'preview_clip_ref', None)
        cam_wgt = getattr(self, 'camera_widget', None)
        if not preview_clip or not cam_wgt:
            return
        area_w, area_h = preview_clip.size
        if area_w <= 1 or area_h <= 1:
            return
        tex = getattr(cam_wgt, 'texture', None)
        if tex and tex.size[0] > 0 and tex.size[1] > 0:
            cam_w, cam_h = float(tex.size[0]), float(tex.size[1])
        else:
            res = getattr(cam_wgt, 'resolution', (640, 480))
            cam_w, cam_h = float(res[0]), float(res[1])

        is_android = platform == 'android'
        upright_ratio = (cam_h / cam_w) if is_android else (cam_w / cam_h)
        area_ratio = area_w / area_h
        if area_ratio > upright_ratio:
            disp_w = area_w
            disp_h = area_w / upright_ratio
        else:
            disp_h = area_h
            disp_w = area_h * upright_ratio

        cam_wgt.size_hint = (None, None)
        cam_wgt.size = (disp_h, disp_w) if is_android else (disp_w, disp_h)
        cam_wgt.center = preview_clip.center
        if getattr(self, 'rotation', None):
            self.rotation.origin = cam_wgt.center

    def _launch_camera_widget(self):
        from kivy.graphics import Color, Line, Rectangle
        from kivy.uix.stencilview import StencilView
        self.scanner_start_time = time.time()
        self._laser_phase = 0.0
        self._laser_dir = 1.0
        try:
            self.camera_widget = Camera(play=True, index=0, resolution=(640, 480), allow_stretch=True, keep_ratio=False)
            with self.camera_widget.canvas.before:
                PushMatrix()
                self.rotation = Rotate(angle=-90 if platform == 'android' else 0, origin=self.camera_widget.center)
            with self.camera_widget.canvas.after:
                PopMatrix()
            self.camera_widget.bind(
                center=lambda i, v: setattr(self.rotation, 'origin', i.center) if getattr(self, 'rotation', None) else None,
                texture=self._update_camera_preview_layout,
            )
        except Exception:
            self.notify('Erreur chargement caméra.', 'error')
            return

        is_single = getattr(self, 'target_scan_field', None) is not None
        root = MDBoxLayout(orientation='vertical')
        cam_area = MDFloatLayout(size_hint_y=1.0 if is_single else 0.48)
        self.cam_area_ref = cam_area

        preview_clip = StencilView(size_hint=(1, 1), pos_hint={'x': 0, 'y': 0})
        self.preview_clip_ref = preview_clip
        preview_clip.add_widget(self.camera_widget)
        preview_clip.bind(pos=self._update_camera_preview_layout, size=self._update_camera_preview_layout)
        cam_area.add_widget(preview_clip)

        with cam_area.canvas.after:
            self._scan_box_color = Color(0.20, 0.83, 0.60, 0.92)
            self._scan_box_line = Line(rounded_rectangle=(0, 0, 100, 100, dp(10)), width=dp(2.2))
            self._laser_color = Color(0.94, 0.27, 0.27, 0.95)
            self._laser_rect = Rectangle(pos=(0, 0), size=(100, dp(2)))
            self._focus_ring_color = Color(0.98, 0.75, 0.14, 0.0)
            self._focus_ring_line = Line(circle=(0, 0, dp(26)), width=dp(2))

        cam_area.bind(pos=self._update_scan_overlay_geometry, size=self._update_scan_overlay_geometry)

        def _on_cam_touch(inst, touch):
            if not inst.collide_point(*touch.pos):
                return False
            for ctrl in (
                getattr(self, 'btn_scan_close', None),
                getattr(self, 'btn_scan_sound', None),
                getattr(self, 'btn_scan_zoom', None),
                getattr(self, 'btn_scan_torch', None),
                getattr(self, 'btn_scan_pause', None),
            ):
                if ctrl and not ctrl.disabled and ctrl.opacity > 0 and ctrl.collide_point(*touch.pos):
                    return False
            self._trigger_tap_to_focus(touch.pos)
            return True

        cam_area.bind(on_touch_down=_on_cam_touch)

        self.btn_scan_close = MDIconButton(
            icon='close', icon_size='28sp',
            md_bg_color=(0.8, 0.15, 0.15, 0.85), theme_text_color='Custom', text_color=(1, 1, 1, 1),
            pos_hint={'top': 0.96, 'right': 0.97}, on_release=self.close_barcode_scanner,
        )
        self.btn_scan_sound = MDIconButton(
            icon='volume-high' if self.scanner_sound_enabled else 'volume-off', icon_size='24sp',
            md_bg_color=(0.15, 0.45, 0.85, 0.85) if self.scanner_sound_enabled else (0.15, 0.15, 0.18, 0.75),
            theme_text_color='Custom', text_color=(1, 1, 1, 1),
            pos_hint={'top': 0.96, 'x': 0.03}, on_release=self._toggle_scanner_sound,
        )
        self.btn_scan_pause = MDIconButton(
            icon='pause', icon_size='24sp',
            md_bg_color=(0.15, 0.15, 0.18, 0.75), theme_text_color='Custom', text_color=(1, 1, 1, 1),
            pos_hint={'y': 0.04, 'right': 0.97}, on_release=self._toggle_camera_pause,
        )
        self.btn_scan_torch = MDIconButton(
            icon='flashlight-off', icon_size='24sp',
            md_bg_color=(0.15, 0.15, 0.18, 0.75), theme_text_color='Custom', text_color=(1, 1, 1, 1),
            pos_hint={'y': 0.04, 'x': 0.03}, opacity=0, disabled=True, on_release=self._toggle_camera_torch,
        )
        self.btn_scan_zoom = MDRaisedButton(
            text='1x', font_size='14sp', elevation=0,
            md_bg_color=(0.15, 0.15, 0.18, 0.75), text_color=(1, 1, 1, 1),
            size_hint=(None, None), size=(dp(56), dp(40)),
            pos_hint={'y': 0.04, 'center_x': 0.5}, opacity=0, disabled=True, on_release=self._toggle_camera_zoom,
        )
        cam_area.add_widget(self.btn_scan_close)
        cam_area.add_widget(self.btn_scan_sound)
        cam_area.add_widget(self.btn_scan_pause)
        cam_area.add_widget(self.btn_scan_torch)
        cam_area.add_widget(self.btn_scan_zoom)
        root.add_widget(cam_area)

        manual_bar = MDCard(orientation='horizontal', size_hint_y=None, height=dp(58), padding=[dp(10), dp(4)], spacing=dp(8), radius=[0], elevation=0, md_bg_color=(0.97, 0.98, 0.99, 1))
        self.scan_manual_field = MDTextField(hint_text='Saisir le code-barres...', mode='rectangle', multiline=False, size_hint_x=0.75, font_size='16sp', pos_hint={'center_y': 0.5})
        self.scan_manual_field.bind(on_text_validate=lambda x: self._handle_manual_scan_submit())
        manual_btn = MDRaisedButton(text='OK', size_hint_x=0.25, size_hint_y=None, height=dp(44), md_bg_color=(0.1, 0.45, 0.85, 1), text_color=(1, 1, 1, 1), pos_hint={'center_y': 0.5}, on_release=lambda x: self._handle_manual_scan_submit())
        manual_bar.add_widget(self.scan_manual_field)
        manual_bar.add_widget(manual_btn)
        root.add_widget(manual_bar)

        if not is_single:
            box = MDCard(orientation='vertical', size_hint_y=0.52, radius=[16, 16, 0, 0], md_bg_color=(1, 1, 1, 1), elevation=0)
            hdr = MDBoxLayout(size_hint_y=None, height=dp(36), padding=[dp(16), 0])
            self.lbl_scan_count = MDLabel(text='Articles scannés : 0', bold=True, theme_text_color='Primary')
            hdr.add_widget(self.lbl_scan_count)
            box.add_widget(hdr)
            scr = MDScrollView()
            self.scan_list_widget = MDList()
            scr.add_widget(self.scan_list_widget)
            box.add_widget(scr)
            box.add_widget(MDRaisedButton(text='TERMINER', font_size='18sp', size_hint=(1, None), height=dp(52), md_bg_color=(0, 0.7, 0.2, 1), elevation=0, on_release=self.finish_continuous_scan))
            root.add_widget(box)

        self.scan_dialog = ModalView(size_hint=(1, 1), auto_dismiss=False, background_color=(0, 0, 0, 1))
        self.scan_dialog.add_widget(root)
        self.scan_dialog.open()
        Clock.schedule_once(lambda dt: [self._update_camera_preview_layout(), self._update_scan_overlay_geometry()], 0.05)
        Clock.schedule_once(lambda dt: [self._update_camera_preview_layout(), self._optimize_camera_params()], 0.45)
        Clock.schedule_once(lambda dt: [self._update_camera_preview_layout(), self._optimize_camera_params()], 1.00)
        self._laser_ev = Clock.schedule_interval(self._animate_laser_step, 1.0 / 30.0)
        self.scan_event = Clock.schedule_interval(self.detect_barcode_frame, 1.0 / 20.0)

    def _update_scan_overlay_geometry(self, *args):
        cam_area = getattr(self, 'cam_area_ref', None)
        if not cam_area or not getattr(self, '_scan_box_line', None):
            return
        x, y = cam_area.pos
        w, h = cam_area.size
        box_w = max(dp(40), w * 0.84)
        box_h = min(max(dp(40), h * 0.58), dp(260))
        box_x = x + (w - box_w) / 2.0
        box_y = y + (h - box_h) / 2.0
        self._scan_box_rect = (box_x, box_y, box_w, box_h)
        self._scan_box_line.rounded_rectangle = (box_x, box_y, box_w, box_h, dp(10))
        if getattr(self, '_laser_rect', None):
            ly = box_y + dp(8) + max(0.0, (box_h - dp(16))) * getattr(self, '_laser_phase', 0.5)
            self._laser_rect.pos = (box_x + dp(8), ly)
            self._laser_rect.size = (max(dp(10), box_w - dp(16)), dp(2))

    def _animate_laser_step(self, dt):
        if getattr(self, 'is_scanner_paused', False) or not getattr(self, '_laser_rect', None):
            return
        rect = getattr(self, '_scan_box_rect', None)
        if not rect:
            return
        box_x, box_y, box_w, box_h = rect
        step = dt * 1.25 * getattr(self, '_laser_dir', 1.0)
        phase = getattr(self, '_laser_phase', 0.0) + step
        if phase >= 1.0:
            phase = 1.0
            self._laser_dir = -1.0
        elif phase <= 0.0:
            phase = 0.0
            self._laser_dir = 1.0
        self._laser_phase = phase
        ly = box_y + dp(8) + max(0.0, (box_h - dp(16))) * phase
        self._laser_rect.pos = (box_x + dp(8), ly)
        self._laser_rect.size = (max(dp(10), box_w - dp(16)), dp(2))

    def _flash_scan_frame(self, is_success: bool):
        if not getattr(self, '_scan_box_color', None):
            return
        if getattr(self, '_flash_reset_ev', None):
            self._flash_reset_ev.cancel()
        self._scan_box_color.rgba = (0.06, 0.72, 0.50, 1.0) if is_success else (0.94, 0.27, 0.27, 1.0)
        self._flash_reset_ev = Clock.schedule_once(
            lambda dt: setattr(self._scan_box_color, 'rgba', (0.20, 0.83, 0.60, 0.92)) if getattr(self, '_scan_box_color', None) else None,
            0.46,
        )

    def close_barcode_scanner(self, *args):
        self._scanner_session_id = getattr(self, '_scanner_session_id', 0) + 1
        for ev_attr in ('scan_event', '_laser_ev', '_flash_reset_ev', '_focus_ring_ev'):
            ev = getattr(self, ev_attr, None)
            if ev is not None:
                ev.cancel()
                setattr(self, ev_attr, None)
        if getattr(self, '_torch_active', False):
            acam = self._get_android_camera()
            if acam is not None:
                try:
                    params = acam.getParameters()
                    if params is not None:
                        params.setFlashMode('off')
                        acam.setParameters(params)
                except Exception:
                    pass
            self._torch_active = False
        if getattr(self, 'camera_widget', None):
            try:
                self.camera_widget.play = False
                cam_core = getattr(self.camera_widget, '_camera', None)
                if cam_core is not None:
                    if hasattr(cam_core, 'stop'):
                        cam_core.stop()
                    acam = getattr(cam_core, '_android_camera', None) or getattr(cam_core, '_camera', None)
                    if acam is not None:
                        try:
                            acam.release()
                        except Exception:
                            pass
                        if hasattr(cam_core, '_android_camera'):
                            cam_core._android_camera = None
                        if hasattr(cam_core, '_camera'):
                            cam_core._camera = None
            except Exception:
                pass
            self.camera_widget = None
        if getattr(self, 'scan_dialog', None):
            self.scan_dialog.dismiss()
            self.scan_dialog = None
        self.cam_area_ref = None
        self.preview_clip_ref = None
        self._scan_box_color = None
        self._scan_box_line = None
        self._laser_color = None
        self._laser_rect = None
        self._focus_ring_color = None
        self._focus_ring_line = None
        self.is_detecting_frame = False
        self.is_scanner_paused = False
        self.temp_scanned_cart = []
        self.target_scan_field = None

    def detect_barcode_frame(self, dt):
        if not getattr(self, 'scan_dialog', None) or not self.scan_dialog.parent:
            return
        if getattr(self, 'is_scanner_paused', False) or getattr(self, 'is_detecting_frame', False):
            return
        if not getattr(self, 'camera_widget', None) or not self.camera_widget.texture:
            return
        if time.time() - getattr(self, 'scanner_start_time', 0) < 0.8:
            return
        try:
            tex = self.camera_widget.texture
            raw_pixels = bytes(tex.pixels)
            tex_size = tuple(tex.size)
            color_fmt = str(getattr(tex, 'colorfmt', 'rgba') or 'rgba').upper()
            if not raw_pixels or tex_size[0] <= 0 or tex_size[1] <= 0:
                return
            self.is_detecting_frame = True
            sid = getattr(self, '_scanner_session_id', 0)
            threading.Thread(target=self._decode_frame_worker, args=(raw_pixels, tex_size, color_fmt, sid), daemon=True).start()
        except Exception:
            self.is_detecting_frame = False

    def _detect_frame_target(self, img: Image.Image) -> tuple:
        vw, vh = img.size
        self._frame_scan_tick = getattr(self, '_frame_scan_tick', 0) + 1
        if vw < 400 or vh < 300 or (self._frame_scan_tick % 4 == 0):
            return img, False
        sx = int(vw * 0.02)
        sy = int(vh * 0.16)
        ex = max(sx + 160, int(vw * 0.98))
        ey = max(sy + 120, int(vh * 0.84))
        return img.crop((sx, sy, min(vw, ex), min(vh, ey))), True

    def _pick_best_detected_code(self, bcs, is_roi: bool, surf_w: int, surf_h: int):
        if not bcs:
            return None
        center_x = surf_w / 2.0
        center_y = surf_h / 2.0
        best = None
        best_dist = float('inf')

        for b in bcs:
            raw_data = getattr(b, 'data', None)
            if not raw_data:
                continue
            rect = getattr(b, 'rect', None)
            if rect and rect.width > 0 and rect.height > 0:
                if rect.width < 20 or rect.height < 10:
                    continue
                if is_roi and surf_w > 0 and surf_h > 0:
                    if rect.left < 4 or rect.top < 4 or (rect.left + rect.width) > (surf_w - 4) or (rect.top + rect.height) > (surf_h - 4):
                        continue
                bx = rect.left + rect.width / 2.0
                by = rect.top + rect.height / 2.0
                dist = math.hypot(bx - center_x, by - center_y)
                if dist < best_dist:
                    best_dist = dist
                    best = b
            elif best is None:
                best = b

        if best is None:
            return None
        code_str = best.data.decode('utf-8', errors='ignore').strip()
        fmt_str = str(getattr(best, 'type', '') or '')
        return (code_str, fmt_str) if code_str else None

    def _decode_frame_worker(self, raw_pixels: bytes, tex_size: tuple, color_fmt: str, session_id: int):
        try:
            if session_id != getattr(self, '_scanner_session_id', 0):
                return
            mode = 'RGB' if color_fmt == 'RGB' else 'RGBA'
            img = Image.frombytes(mode=mode, size=tex_size, data=raw_pixels).convert('L')
            if platform == 'android':
                img = img.transpose(Image.ROTATE_270)
            target_img, is_roi = self._detect_frame_target(img)
            symbols = self._get_retail_zbar_symbols()
            bcs = zbar_decode(target_img, symbols=symbols) if symbols else zbar_decode(target_img)
            picked = self._pick_best_detected_code(bcs, is_roi, target_img.size[0], target_img.size[1])
            if picked and session_id == getattr(self, '_scanner_session_id', 0):
                code_str, fmt_str = picked
                Clock.schedule_once(lambda dt, c=code_str, f=fmt_str, sid=session_id: self._on_barcode_detected(c, f, sid), 0)
        except Exception:
            pass
        finally:
            self.is_detecting_frame = False





    def _evaluate_retail_code(self, code: str, format_str: str) -> dict:
        if not code or len(code) < 2 or len(code) > 64:
            return {'valid': False, 'checksum_ok': False, 'is_qr': False, 'is_weak_1d': False}
        fmt = str(format_str or '').lower()
        if 'qr' in fmt:
            return {'valid': True, 'checksum_ok': True, 'is_qr': True, 'is_weak_1d': False}
        if fmt in ('upce', 'upc_e', 'upc-e'):
            return {'valid': code.isdigit() and len(code) in (6, 8), 'checksum_ok': False, 'is_qr': False, 'is_weak_1d': False}

        is_weak_1d = any(k in fmt for k in ('i25', 'itf', 'codabar', 'code39', 'code_39'))
        if code.isdigit() and len(code) in (8, 12, 13, 14):
            digits = [int(ch) for ch in code]
            check = digits[-1]
            total = 0
            pos_from_right = 1
            for i in range(len(digits) - 2, -1, -1):
                total += digits[i] * (3 if pos_from_right % 2 == 1 else 1)
                pos_from_right += 1
            expected = (10 - (total % 10)) % 10
            matched = check == expected
            if 'ean' in fmt or 'upc' in fmt:
                return {'valid': matched, 'checksum_ok': matched, 'is_qr': False, 'is_weak_1d': False}
            return {'valid': True, 'checksum_ok': matched and len(code) in (8, 12, 13), 'is_qr': False, 'is_weak_1d': is_weak_1d}
        return {'valid': True, 'checksum_ok': False, 'is_qr': False, 'is_weak_1d': is_weak_1d}

    def _on_barcode_detected(self, raw_code: str, format_str: str, session_id: int):
        if session_id != getattr(self, '_scanner_session_id', 0):
            return
        if not getattr(self, 'scan_dialog', None) or not self.scan_dialog.parent:
            return
        if getattr(self, 'is_scanner_paused', False):
            return

        clean_code = re.sub(r'[\r\n\t\x00-\x1f\x7f]', '', str(raw_code or '')).strip()
        if not clean_code:
            return

        check = self._evaluate_retail_code(clean_code, format_str)
        if not check['valid']:
            return

        now = time.time()
        if clean_code == getattr(self, 'potential_code', None) and (now - getattr(self, '_consensus_ts', 0.0)) <= 0.95:
            self.consecutive_frames = getattr(self, 'consecutive_frames', 0) + 1
        else:
            self.potential_code = clean_code
            self.consecutive_frames = 1
        self._consensus_ts = now

        required_hits = 3
        if check['is_qr']:
            required_hits = 2
        elif check['is_weak_1d'] or (not check['checksum_ok'] and len(clean_code) <= 5):
            required_hits = 4
        if self.consecutive_frames < required_hits:
            return

        cooldown = 1.5
        recent_map = getattr(self, '_recent_scan_times', {})
        last_code_ts = recent_map.get(clean_code, 0.0)
        if (clean_code == getattr(self, '_last_scanned_code', '') and (now - getattr(self, 'last_scan_time', 0.0)) < cooldown) or (now - last_code_ts) < cooldown:
            return
        if (now - getattr(self, '_last_any_scan_time', 0.0)) < 0.4:
            return

        self._last_scanned_code = clean_code
        self.last_scan_time = now
        self._last_any_scan_time = now
        recent_map[clean_code] = now
        if len(recent_map) > 100:
            for k in [k for k, v in recent_map.items() if (now - v) > 10.0]:
                del recent_map[k]
        self._recent_scan_times = recent_map
        self.potential_code = None
        self.consecutive_frames = 0

        self.process_continuous_scan(clean_code)

    def _handle_manual_scan_submit(self):
        if not getattr(self, 'scan_manual_field', None):
            return
        code = re.sub(r'[\r\n\t\x00-\x1f\x7f]', '', str(self.scan_manual_field.text or '')).strip()
        if not code:
            return
        self.scan_manual_field.text = ''
        self.scan_manual_field.focus = False
        self.process_continuous_scan(code)

    def process_continuous_scan(self, code: str):
        if not getattr(self, 'scan_dialog', None) or not self.scan_dialog.parent:
            return
        sound_on = getattr(self, 'scanner_sound_enabled', True)
        if getattr(self, 'target_scan_field', None) is not None:
            self.target_scan_field.text = code
            self._flash_scan_frame(True)
            if sound_on:
                self.play_sound('success')
            self.close_barcode_scanner()
            return
        prod = self.db.get_product_by_barcode(code)
        if not prod:
            self._flash_scan_frame(False)
            if sound_on:
                self.play_sound('error')
            self.notify(f'Code-barres introuvable : {code}', 'error')
            return
        if any(it['id'] == prod['id'] for it in self.temp_scanned_cart):
            self._flash_scan_frame(False)
            if sound_on:
                self.play_sound('duplicate')
            self.notify(f"Déjà scanné : {self.fix_text(prod['name'])}", 'warning')
            return
        ok_stk, msg_stk = self._check_stock_availability(prod, Decimal('1'))
        if not ok_stk:
            self._flash_scan_frame(False)
            if sound_on:
                self.play_sound('error')
            self.notify(msg_stk.replace('\n', ' '), 'error')
            return
        self.temp_scanned_cart.append(prod)
        self._refresh_scan_list_ui()
        self._flash_scan_frame(True)
        if sound_on:
            self.play_sound('success')

    def _refresh_scan_list_ui(self):
        self.scan_list_widget.clear_widgets()
        self.lbl_scan_count.text = f'Articles scannés : {len(self.temp_scanned_cart)}'
        doc_type = AppConstants.get_doc_type_for_mode(self.current_mode)
        is_qty_only = doc_type in AppConstants.QTY_ONLY_DOC_TYPES
        for prod in reversed(self.temp_scanned_cart):
            price, _, from_hist = self._resolve_product_unit_price(prod, doc_type)
            card = MDCard(orientation='horizontal', size_hint_y=None, height=dp(72), padding=[dp(15), 0, 0, 0], radius=[0], elevation=0)
            tb = MDBoxLayout(orientation='vertical', pos_hint={'center_y': 0.5}, adaptive_height=True)
            tb.add_widget(MDLabel(text=self.fix_text(prod['name']), font_style='Subtitle1', max_lines=2, adaptive_height=True))
            if is_qty_only:
                sub_txt = f"Stock Magasin : {FinancialEngine.format_qty(prod.get('stock_store', 0))}"
            else:
                sub_txt = f'Prix : {FinancialEngine.format_money(price)}' + (' (Historique)' if from_hist else '')
            tb.add_widget(MDLabel(text=sub_txt, font_style='Caption', theme_text_color='Secondary', bold=True, adaptive_height=True))
            card.add_widget(tb)
            card.add_widget(MDIconButton(icon='delete', theme_text_color='Custom', text_color=(0.9, 0, 0, 1), pos_hint={'center_y': 0.5}, on_release=lambda x, p=prod: [self.temp_scanned_cart.remove(p), self._refresh_scan_list_ui()]))
            self.scan_list_widget.add_widget(card)
            self.scan_list_widget.add_widget(MDSeparator(height=dp(1), color=(0.95, 0.95, 0.95, 1)))

    def finish_continuous_scan(self, instance):
        items = list(self.temp_scanned_cart)
        self.close_barcode_scanner()
        doc_type = AppConstants.get_doc_type_for_mode(self.current_mode)
        added = 0
        for prod in items:
            price, _, _ = self._resolve_product_unit_price(prod, doc_type)
            ok_stk, _ = self._check_stock_availability(prod, Decimal('1'))
            if not ok_stk:
                continue
            existing = next((it for it in self.cart if it['id'] == prod['id']), None)
            if existing:
                existing['qty'] = float(FinancialEngine.quantize_qty(FinancialEngine.to_decimal(existing['qty']) + Decimal('1')))
                existing['price'] = float(price)
            else:
                self.cart.append({
                    'id': prod['id'],
                    'name': prod['name'],
                    'price': float(price),
                    'qty': 1.0,
                    'tva': float(prod.get('tva', 0)) if doc_type in AppConstants.INVOICE_TAX_DOC_TYPES else 0.0,
                    'discount_amount': 0.0,
                    'discount_type': 'fixed',
                })
            added += 1
        self.update_cart_button()
        if added > 0:
            self.notify(f'{added} article(s) ajouté(s) au panier.', 'success')

    # --------------------------------------------------------------------------
    # 7.13 PARAMÈTRES, MAGASIN (AVEC LOGO), SÉCURITÉ VENDEUR, EXCEL & SAUVEGARDE ZIP
    # --------------------------------------------------------------------------
    def open_settings_menu(self, *args):
        if getattr(self, 'settings_menu_dialog', None):
            try:
                self.settings_menu_dialog.dismiss()
            except Exception:
                pass
            self.settings_menu_dialog = None

        wrapper = MDBoxLayout(
            orientation='vertical',
            spacing=dp(10),
            size_hint_y=None,
            height=min(dp(440) if self.is_seller_mode else dp(595), Window.height * 0.82),
            padding=[0, dp(2), 0, 0],
        )
        scroll_view = MDScrollView(size_hint_y=1, bar_width=dp(3))
        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(12),
            size_hint_y=None,
            adaptive_height=True,
            padding=[0, dp(2), dp(2), dp(6)],
        )

        def _make_section_card():
            return MDCard(
                orientation='vertical',
                padding=dp(12),
                spacing=dp(8),
                size_hint_y=None,
                adaptive_height=True,
                radius=[14],
                elevation=0,
                md_bg_color=(0.97, 0.98, 0.99, 1),
                line_color=(0.88, 0.90, 0.94, 1),
                line_width=1,
            )

        def _add_card_header(parent_card, icon_name: str, title_text: str):
            hdr = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint_y=None, height=dp(24))
            hdr.add_widget(MDIcon(icon=icon_name, theme_text_color='Custom', text_color=(0.09, 0.44, 0.84, 1), font_size='19sp', size_hint=(None, None), size=(dp(20), dp(20)), pos_hint={'center_y': 0.5}))
            hdr.add_widget(MDLabel(text=title_text, font_style='Caption', bold=True, theme_text_color='Custom', text_color=(0.09, 0.44, 0.84, 1), size_hint_x=1, pos_hint={'center_y': 0.5}))
            parent_card.add_widget(hdr)

        def _add_action_row(parent_card, title_txt: str, sub_txt: str, icon_name: str, icon_col, bg_col, cb, right_icon='chevron-right', title_col=None):
            row_card = MDCard(
                orientation='horizontal',
                padding=[dp(10), dp(8), dp(10), dp(8)],
                spacing=dp(10),
                size_hint_y=None,
                height=dp(58),
                radius=[11],
                elevation=0,
                md_bg_color=bg_col,
                line_color=(0.89, 0.91, 0.94, 1),
                line_width=1,
                ripple_behavior=True,
                on_release=cb,
            )
            row_card.add_widget(self._ui_icon_pill(icon_name, (icon_col[0], icon_col[1], icon_col[2], 0.12), icon_col, size_dp=36, icon_size='20sp'))

            t_box = MDBoxLayout(orientation='vertical', size_hint_x=1, size_hint_y=None, adaptive_height=True, pos_hint={'center_y': 0.5}, spacing=dp(2))
            t_box.add_widget(self._ui_wrap_label(title_txt, font_style='Subtitle2', bold=True, theme_color='Custom' if title_col else 'Primary', text_color=title_col, min_h=20))
            t_box.add_widget(self._ui_wrap_label(sub_txt, font_style='Caption', theme_color='Secondary', min_h=16))
            t_box.bind(height=lambda inst, val: setattr(row_card, 'height', max(dp(58), val + dp(16))))
            row_card.add_widget(t_box)
            row_card.add_widget(MDIcon(icon=right_icon, theme_text_color='Hint', font_size='21sp', size_hint=(None, None), size=(dp(22), dp(22)), pos_hint={'center_y': 0.5}))
            parent_card.add_widget(row_card)

        def _add_checkbox_row(parent_card, label_txt: str, setting_key: str, default_val: bool = False):
            row = MDBoxLayout(orientation='horizontal', size_hint_y=None, adaptive_height=True, padding=[dp(6), dp(2), 0, dp(2)], spacing=dp(8))
            lbl = self._ui_wrap_label(label_txt, font_style='Body2', theme_color='Primary', min_h=24)
            lbl.pos_hint = {'center_y': 0.5}
            chk = MDCheckbox(active=self.db.get_bool_setting(setting_key, default_val), size_hint=(None, None), size=(dp(40), dp(40)), pos_hint={'center_y': 0.5})
            chk.bind(active=lambda inst, val, k=setting_key: self.db.set_setting(k, str(val)))
            row.add_widget(lbl)
            row.add_widget(chk)
            parent_card.add_widget(row)

        # 1. Carte Statut de Licence
        status, days_left = self.check_license_validity()
        dev_id = self.get_device_id()
        if status == 'ACTIVATED':
            activ_key = hashlib.sha256(f'magpro_mobile_v6_{dev_id}_secure_key'.encode()).hexdigest()
            lic_card = MDCard(
                orientation='vertical',
                padding=dp(14),
                spacing=dp(6),
                size_hint_y=None,
                adaptive_height=True,
                md_bg_color=(0.93, 0.98, 0.94, 1),
                line_color=(0.72, 0.89, 0.76, 1),
                line_width=1.1,
                radius=[14],
                elevation=0,
                ripple_behavior=True,
                on_release=lambda inst: [Clipboard.copy(activ_key), self.notify("Clé d'activation copiée", 'success')],
            )
            top_row = MDBoxLayout(orientation='horizontal', adaptive_height=True, spacing=dp(10))
            top_row.add_widget(MDIcon(icon='shield-check', theme_text_color='Custom', text_color=(0.04, 0.62, 0.24, 1), font_size='24sp', size_hint=(None, None), size=(dp(26), dp(26)), pos_hint={'center_y': 0.5}))
            top_row.add_widget(self._ui_wrap_label('Licence Activée', font_style='Subtitle1', bold=True, theme_color='Custom', text_color=(0.04, 0.56, 0.22, 1), min_h=22))
            top_row.add_widget(MDIcon(icon='content-copy', theme_text_color='Custom', text_color=(0.04, 0.56, 0.22, 1), font_size='18sp', size_hint=(None, None), size=(dp(20), dp(20)), pos_hint={'center_y': 0.5}))
            lic_card.add_widget(top_row)
            lic_card.add_widget(self._ui_wrap_label("Clé d'activation (appuyez pour copier) :", font_style='Caption', theme_color='Secondary', min_h=16))
            lic_card.add_widget(self._ui_wrap_label(activ_key, font_style='Caption', bold=True, theme_color='Primary', min_h=18, font_name='Roboto'))
        else:
            lic_card = MDCard(
                orientation='vertical',
                padding=dp(14),
                spacing=dp(6),
                size_hint_y=None,
                adaptive_height=True,
                md_bg_color=(1, 0.96, 0.90, 1),
                line_color=(0.94, 0.80, 0.62, 1),
                line_width=1.1,
                radius=[14],
                elevation=0,
                ripple_behavior=True,
                on_release=lambda inst: [Clipboard.copy(dev_id), self.notify('ID copié', 'success')],
            )
            top_row = MDBoxLayout(orientation='horizontal', adaptive_height=True, spacing=dp(10))
            top_row.add_widget(MDIcon(icon='shield-alert', theme_text_color='Custom', text_color=(0.88, 0.45, 0, 1), font_size='24sp', size_hint=(None, None), size=(dp(26), dp(26)), pos_hint={'center_y': 0.5}))
            top_row.add_widget(self._ui_wrap_label(f'Essai gratuit : {days_left} jour(s) restant(s)', font_style='Subtitle1', bold=True, theme_color='Custom', text_color=(0.85, 0.42, 0, 1), min_h=22))
            top_row.add_widget(MDIcon(icon='content-copy', theme_text_color='Custom', text_color=(0.85, 0.42, 0, 1), font_size='18sp', size_hint=(None, None), size=(dp(20), dp(20)), pos_hint={'center_y': 0.5}))
            lic_card.add_widget(top_row)
            lic_card.add_widget(self._ui_wrap_label("ID de l'appareil (appuyez pour copier) :", font_style='Caption', theme_color='Secondary', min_h=16))
            lic_card.add_widget(self._ui_wrap_label(dev_id, font_style='Caption', bold=True, theme_color='Primary', min_h=18, font_name='Roboto'))
        content.add_widget(lic_card)

        if not self.is_seller_mode:
            # 2. Carte Magasin & Règles
            store_card = _make_section_card()
            _add_card_header(store_card, 'store-cog', 'MAGASIN & RÈGLES')
            store_name = self.db.get_setting('store_name', '') or AppConstants.APP_NAME
            _add_action_row(store_card, 'Informations Magasin', self.fix_text(f'Actuel : {store_name}'), 'store', (0.09, 0.44, 0.84, 1), (0.94, 0.97, 1, 1), lambda x: self.show_store_settings_dialog())
            _add_checkbox_row(store_card, 'Autoriser la vente en stock négatif', 'allow_negative_stock', True)
            store_card.add_widget(MDSeparator(height=dp(1), color=(0.9, 0.91, 0.94, 1)))
            _add_checkbox_row(store_card, 'Avertir si prix vente < prix achat', 'warn_price_below_cost', True)
            content.add_widget(store_card)

            # 3. Carte Données & Sauvegarde
            data_card = _make_section_card()
            _add_card_header(data_card, 'database-cog', 'DONNÉES & SAUVEGARDE')
            _add_action_row(data_card, 'Exporter Produits', 'Direct Excel (.xlsx)', 'file-export', (0.06, 0.58, 0.25, 1), (1, 1, 1, 1), lambda x: self.perform_export())
            _add_action_row(data_card, 'Importer Produits', 'Depuis Excel (.xlsx)', 'file-import', (0.09, 0.44, 0.84, 1), (1, 1, 1, 1), lambda x: [self.settings_menu_dialog.dismiss(), self.import_data_dialog()])
            _add_action_row(data_card, 'Sauvegarde Locale', 'Archive complète (.zip)', 'database-export', (0.45, 0.2, 0.72, 1), (1, 1, 1, 1), lambda x: self.perform_local_backup('manual'))
            _add_action_row(data_card, 'Partager Sauvegarde', 'Bluetooth / Quick Share (.zip)', 'cloud-upload', (0.0, 0.55, 0.65, 1), (1, 1, 1, 1), lambda x: self.share_database_file())
            _add_action_row(data_card, 'Restaurer', 'Depuis une archive (.zip / .db)', 'backup-restore', (0.85, 0.45, 0.0, 1), (1, 0.97, 0.92, 1), lambda x: [self.settings_menu_dialog.dismiss(), self.show_restore_dialog()])
            content.add_widget(data_card)

            # 4. Carte Imprimante & PDF
            printer_card = _make_section_card()
            _add_card_header(printer_card, 'printer', 'IMPRIMANTE & PDF')
            p_name = self.db.get_setting('printer_name', 'Non configurée') or 'Non configurée'
            _add_action_row(printer_card, 'Configurer Imprimante', f'Actuelle : {p_name}', 'printer-wireless', (0.09, 0.44, 0.84, 1), (1, 1, 1, 1), lambda x: [self.settings_menu_dialog.dismiss(), self.open_bluetooth_selector(x)])
            _add_action_row(printer_card, "Oublier l'imprimante", "Dissocier l'appareil actuel", 'printer-off', (0.82, 0.14, 0.14, 1), (1, 0.95, 0.95, 1), self.clear_printer_selection, title_col=(0.82, 0.14, 0.14, 1))
            _add_checkbox_row(printer_card, 'Impression Auto (Bluetooth 80 mm)', 'printer_auto', True)
            printer_card.add_widget(MDSeparator(height=dp(1), color=(0.9, 0.91, 0.94, 1)))
            _add_checkbox_row(printer_card, 'Afficher le solde sur PDF (BV / BA)', 'show_balance_in_pdf', False)
            content.add_widget(printer_card)

        # 5. Carte Sécurité
        sec_card = _make_section_card()
        _add_card_header(sec_card, 'shield-account', 'SÉCURITÉ')
        mode_lbl = 'Mode Vendeur (Restreint)' if self.is_seller_mode else 'Mode Administrateur (Complet)'
        _add_action_row(
            sec_card,
            'Mode Admin / Vendeur',
            f'Actif : {mode_lbl}',
            'shield-key',
            (0.28, 0.32, 0.78, 1),
            (1, 1, 1, 1),
            lambda x: [self.settings_menu_dialog.dismiss(), Clock.schedule_once(lambda dt: self.open_seller_auth_dialog(), 0.12)],
        )
        content.add_widget(sec_card)

        # 6. Carte Mise à jour & Version
        update_card = MDCard(
            orientation='horizontal',
            padding=[dp(12), dp(10)],
            spacing=dp(12),
            size_hint_y=None,
            height=dp(60),
            radius=[14],
            elevation=0,
            md_bg_color=(0.92, 0.97, 0.97, 1),
            line_color=(0.76, 0.89, 0.89, 1),
            line_width=1,
            ripple_behavior=True,
            on_release=lambda x: [self.settings_menu_dialog.dismiss(), webbrowser.open('https://mobile.magpro-soft.com/')],
        )
        update_card.add_widget(self._ui_icon_pill('cloud-download', (0.08, 0.55, 0.55, 0.14), (0.08, 0.55, 0.55, 1), size_dp=38, icon_size='22sp'))

        upd_txt = MDBoxLayout(orientation='vertical', size_hint_x=1, size_hint_y=None, adaptive_height=True, pos_hint={'center_y': 0.5}, spacing=dp(2))
        upd_txt.add_widget(self._ui_wrap_label("Mise à jour de l'application", font_style='Subtitle2', bold=True, theme_color='Primary', min_h=20))
        upd_txt.add_widget(self._ui_wrap_label(f'Version actuelle : {AppConstants.APP_VERSION} • © {AppConstants.COPYRIGHT_YEAR}', font_style='Caption', theme_color='Secondary', min_h=16))
        upd_txt.bind(height=lambda inst, val: setattr(update_card, 'height', max(dp(60), val + dp(20))))
        update_card.add_widget(upd_txt)
        update_card.add_widget(MDIcon(icon='open-in-new', theme_text_color='Hint', font_size='20sp', size_hint=(None, None), size=(dp(22), dp(22)), pos_hint={'center_y': 0.5}))
        content.add_widget(update_card)

        scroll_view.add_widget(content)
        wrapper.add_widget(scroll_view)
        wrapper.add_widget(self._ui_card_btn('FERMER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), height=46, on_release=lambda *a: self.settings_menu_dialog.dismiss() if getattr(self, 'settings_menu_dialog', None) else None))

        self.settings_menu_dialog = MDDialog(
            title='Paramètres & Configuration',
            type='custom',
            content_cls=wrapper,
            radius=[18, 18, 18, 18],
            buttons=[],
            size_hint=(0.96, None),
        )
        self._fit_dialog(self.settings_menu_dialog)
        self.settings_menu_dialog.open()

    def show_store_settings_dialog(self):
        if getattr(self, 'settings_menu_dialog', None):
            self.settings_menu_dialog.dismiss()

        current_logo = self._resolve_store_logo_path()
        self._temp_logo_path = None
        self._remove_logo_order = False

        wrapper = MDBoxLayout(
            orientation='vertical',
            spacing=dp(10),
            size_hint_y=None,
            height=min(dp(595), Window.height * 0.82),
            padding=[0, dp(2), 0, 0],
        )
        scroll = MDScrollView(size_hint_y=1, bar_width=dp(3))
        main_box = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(10), padding=[0, dp(2), dp(2), dp(6)])

        # --- Section Logo du Magasin ---
        card_logo = self._ui_section_card((0.95, 0.97, 1, 1), (0.82, 0.89, 0.97, 1), 'image-area', 'Logo du Magasin (PDF A4 & 80mm)')
        logo_row = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint_y=None, height=dp(58))
        logo_thumb_card = MDCard(size_hint=(None, None), size=(dp(56), dp(56)), radius=[12], elevation=0, md_bg_color=(1, 1, 1, 1), line_color=(0.84, 0.88, 0.94, 1), line_width=1, pos_hint={'center_y': 0.5})
        logo_fl = MDFloatLayout()
        logo_placeholder_ico = MDIcon(icon='storefront-outline', theme_text_color='Hint', font_size='28sp', size_hint=(None, None), size=(dp(30), dp(30)), halign='center', valign='middle', pos_hint={'center_x': 0.5, 'center_y': 0.5}, opacity=0 if current_logo else 1)
        logo_preview_img = FitImage(source=current_logo if current_logo else '', radius=[11], mipmap=True, opacity=1 if current_logo else 0, pos_hint={'center_x': 0.5, 'center_y': 0.5})
        logo_fl.add_widget(logo_placeholder_ico)
        logo_fl.add_widget(logo_preview_img)
        logo_thumb_card.add_widget(logo_fl)
        logo_row.add_widget(logo_thumb_card)

        lbl_logo_status = MDLabel(
            text='Logo actif' if current_logo else 'Aucun logo',
            font_style='Caption',
            bold=True,
            theme_text_color='Custom',
            text_color=(0.04, 0.58, 0.22, 1) if current_logo else (0.45, 0.48, 0.54, 1),
            size_hint_x=1,
            valign='middle',
            pos_hint={'center_y': 0.5},
        )
        logo_row.add_widget(lbl_logo_status)

        def _on_logo_picked(selection):
            if selection:
                self._temp_logo_path = selection[0]
                self._remove_logo_order = False
                logo_preview_img.source = selection[0]
                logo_preview_img.opacity = 1
                logo_placeholder_ico.opacity = 0
                lbl_logo_status.text = os.path.basename(selection[0])
                lbl_logo_status.text_color = (0.04, 0.58, 0.22, 1)
                self.notify('Logo sélectionné.', 'success')

        def _on_logo_removed(*args):
            self._temp_logo_path = None
            self._remove_logo_order = True
            logo_preview_img.source = ''
            logo_preview_img.opacity = 0
            logo_placeholder_ico.opacity = 1
            lbl_logo_status.text = 'Aucun logo'
            lbl_logo_status.text_color = (0.45, 0.48, 0.54, 1)

        logo_row.add_widget(
            self._ui_icon_pill(
                'trash-can-outline',
                (1, 0.93, 0.93, 1),
                (0.82, 0.14, 0.14, 1),
                on_release=_on_logo_removed,
                border_color=(0.94, 0.78, 0.78, 1),
                size_dp=38,
            )
        )
        btn_pick_logo = self._ui_card_btn('Choisir', (0.09, 0.44, 0.84, 1), height=38, radius=10, font_size='12.5sp', on_release=lambda x: self._open_image_picker(_on_logo_picked))
        btn_pick_logo.size_hint_x = None
        btn_pick_logo.width = dp(80)
        btn_pick_logo.pos_hint = {'center_y': 0.5}
        logo_row.add_widget(btn_pick_logo)
        card_logo.add_widget(logo_row)
        main_box.add_widget(card_logo)

        fields = {}
        card_id = self._ui_section_card((1, 1, 1, 1), (0.88, 0.90, 0.94, 1), 'store-outline', 'Identité du Magasin')
        for key, label, ic in (('store_name', 'Nom du Magasin', 'store'), ('store_activity', 'Activité', 'briefcase')):
            tf = SmartTextField(text=self.db.get_setting(key, ''), hint_text=label, icon_right=ic)
            fields[key] = tf
            card_id.add_widget(tf)
        main_box.add_widget(card_id)

        card_ct = self._ui_section_card((0.96, 0.98, 1, 1), (0.82, 0.89, 0.97, 1), 'card-account-phone-outline', 'Coordonnées')
        for key, label, ic in (('store_phone', 'Téléphone', 'phone'), ('store_address', 'Adresse', 'map-marker'), ('store_email', 'Email', 'email')):
            tf = SmartTextField(text=self.db.get_setting(key, ''), hint_text=label, icon_right=ic)
            fields[key] = tf
            card_ct.add_widget(tf)
        main_box.add_widget(card_ct)

        card_fisc = self._ui_section_card((0.97, 0.97, 0.98, 1), (0.88, 0.89, 0.92, 1), 'file-document-multiple-outline', 'Informations Fiscales & Bancaires')
        for key, label in (('store_rc', 'N° RC'), ('store_nif', 'N.I.F'), ('store_nis', 'N.I.S'), ('store_nai', 'N.A.I'), ('store_rib', 'R.I.B')):
            tf = SmartTextField(text=self.db.get_setting(key, ''), hint_text=label)
            fields[key] = tf
            card_fisc.add_widget(tf)
        main_box.add_widget(card_fisc)

        scroll.add_widget(main_box)
        wrapper.add_widget(scroll)

        def _save(x):
            for k, tf in fields.items():
                self.db.set_setting(k, tf.get_value().strip())

            if self._temp_logo_path:
                if current_logo and os.path.exists(current_logo):
                    try:
                        os.remove(current_logo)
                    except Exception:
                        pass
                saved_logo = self._save_store_logo_local(self._temp_logo_path)
                self.db.set_setting('store_logo', saved_logo)
            elif self._remove_logo_order:
                if current_logo and os.path.exists(current_logo):
                    try:
                        os.remove(current_logo)
                    except Exception:
                        pass
                self.db.set_setting('store_logo', '')

            self.notify('Informations du magasin enregistrées.', 'success')
            self.store_settings_dialog.dismiss()
            self.open_settings_menu()

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(48))
        btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.35, on_release=lambda *a: [self.store_settings_dialog.dismiss(), self.open_settings_menu()]))
        btns.add_widget(self._ui_card_btn('ENREGISTRER', (0.0, 0.62, 0.25, 1), icon='check-circle-outline', size_hint_x=0.65, font_size='15sp', elevation=1, on_release=_save))
        wrapper.add_widget(btns)

        self.store_settings_dialog = MDDialog(
            title='Informations du Magasin',
            type='custom',
            content_cls=wrapper,
            size_hint=(0.96, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.store_settings_dialog)
        self.store_settings_dialog.open()

    def open_seller_auth_dialog(self, _retry: int = 0):
        if getattr(self, '_mode_switch_in_progress', False):
            return
        if getattr(self, 'settings_menu_dialog', None) and self.settings_menu_dialog.parent and _retry < 5:
            Clock.schedule_once(lambda dt: self.open_seller_auth_dialog(_retry + 1), 0.10)
            return

        for d_attr in ('auth_dialog', 'toggle_dialog'):
            dlg = getattr(self, d_attr, None)
            if dlg:
                try:
                    dlg.dismiss()
                except Exception:
                    pass
                setattr(self, d_attr, None)
        Window.release_all_keyboards()

        stored = (self.db.get_setting('seller_password', '') or '').strip()
        if not stored:
            self._show_seller_pin_required_dialog()
            return

        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(12),
            size_hint_y=None,
            adaptive_height=True,
            padding=[0, dp(4), 0, 0],
        )
        content.add_widget(
            self._ui_wrap_label(
                'Saisissez votre code PIN pour basculer entre le mode Administrateur et le mode Vendeur.',
                font_style='Caption',
                halign='center',
                theme_color='Secondary',
                min_h=18,
            )
        )

        pin_card = MDCard(
            orientation='horizontal',
            size_hint=(1, None),
            height=dp(60),
            radius=[12],
            elevation=0,
            md_bg_color=(0.96, 0.97, 0.99, 1),
            line_color=(0.82, 0.88, 0.96, 1),
            line_width=1.1,
            padding=[dp(12), dp(2), dp(6), dp(2)],
            spacing=dp(8),
        )
        pin_card.add_widget(
            MDIcon(
                icon='shield-key-outline',
                theme_text_color='Custom',
                text_color=(0.28, 0.32, 0.78, 1),
                font_size='22sp',
                size_hint_x=None,
                width=dp(24),
                pos_hint={'center_y': 0.5},
            )
        )
        self.seller_pass_field = MDTextField(
            hint_text='Code PIN',
            password=True,
            halign='center',
            font_size='20sp',
            mode='line',
            active_line=False,
            line_color_normal=(0, 0, 0, 0),
            line_color_focus=(0, 0, 0, 0),
            pos_hint={'center_y': 0.5},
        )
        pin_card.add_widget(self.seller_pass_field)
        pin_card.add_widget(
            MDIconButton(
                icon='eye-off-outline',
                theme_text_color='Secondary',
                icon_size='20sp',
                pos_hint={'center_y': 0.5},
                on_release=lambda b: [
                    setattr(self.seller_pass_field, 'password', not self.seller_pass_field.password),
                    setattr(b, 'icon', 'eye-off-outline' if self.seller_pass_field.password else 'eye-outline'),
                ],
            )
        )
        content.add_widget(pin_card)

        btn = self._ui_card_btn(
            'VALIDER',
            (0.09, 0.44, 0.84, 1),
            icon='check-circle-outline',
            size_hint_x=0.60,
            font_size='13.5sp',
            elevation=1,
        )
        btn.bind(on_release=lambda x: self._verify_seller_pin(stored))
        self.seller_pass_field.bind(on_text_validate=lambda x: self._verify_seller_pin(stored))

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(48))
        btns.add_widget(
            self._ui_card_btn(
                'ANNULER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                size_hint_x=0.40,
                font_size='13sp',
                on_release=lambda x: self.auth_dialog.dismiss() if getattr(self, 'auth_dialog', None) else None,
            )
        )
        btns.add_widget(btn)
        content.add_widget(btns)

        self.auth_dialog = MDDialog(
            title='Accès Admin',
            type='custom',
            content_cls=content,
            size_hint=(0.92, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.auth_dialog, top_pad=16, side_pad=14, bottom_pad=14, spacer_top=8)
        self.auth_dialog.open()

    def _show_seller_pin_required_dialog(self):
        if getattr(self, 'auth_dialog', None):
            try:
                self.auth_dialog.dismiss()
            except Exception:
                pass
            self.auth_dialog = None

        content = MDBoxLayout(
            orientation='vertical',
            size_hint_y=None,
            adaptive_height=True,
            spacing=dp(12),
            padding=[0, dp(2), 0, dp(2)],
        )
        badge = self._ui_icon_pill(
            'shield-lock-outline',
            (0.90, 0.95, 1, 1),
            (0.09, 0.44, 0.84, 1),
            size_dp=58,
            radius=29,
            icon_size='30sp',
        )
        badge.pos_hint = {'center_x': 0.5}
        content.add_widget(badge)

        content.add_widget(
            MDLabel(
                text='Code PIN requis',
                halign='center',
                font_style='H6',
                bold=True,
                theme_text_color='Primary',
                size_hint_y=None,
                height=dp(28),
            )
        )

        msg_card = MDCard(
            orientation='vertical',
            size_hint=(1, None),
            adaptive_height=True,
            radius=[12],
            elevation=0,
            md_bg_color=(0.95, 0.97, 1, 1),
            line_color=(0.82, 0.88, 0.96, 1),
            line_width=1,
            padding=[dp(14), dp(12)],
        )
        msg_card.add_widget(
            self._ui_wrap_label(
                "Aucun code PIN n'est encore configuré.\n\nPour sécuriser le basculement entre le mode Administrateur et le mode Vendeur, vous devez d'abord créer un code PIN.",
                font_style='Body2',
                halign='center',
                bold=True,
                theme_color='Custom',
                text_color=(0.16, 0.22, 0.32, 1),
                min_h=22,
            )
        )
        content.add_widget(msg_card)

        def _go_create(*args):
            if getattr(self, 'auth_dialog', None):
                self.auth_dialog.dismiss()
            Clock.schedule_once(lambda dt: self._open_create_seller_pin_dialog(), 0.24)

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint=(1, None), height=dp(48))
        btns.add_widget(
            self._ui_card_btn(
                'ANNULER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                size_hint_x=0.38,
                font_size='12.5sp',
                on_release=lambda x: self.auth_dialog.dismiss() if getattr(self, 'auth_dialog', None) else None,
            )
        )
        btns.add_widget(
            self._ui_card_btn(
                'CRÉER UN PIN',
                (0.09, 0.44, 0.84, 1),
                icon='shield-plus-outline',
                size_hint_x=0.62,
                font_size='12.5sp',
                elevation=1,
                on_release=_go_create,
            )
        )
        content.add_widget(btns)

        self.auth_dialog = MDDialog(
            title='',
            type='custom',
            content_cls=content,
            size_hint=(0.92, None),
            radius=[20, 20, 20, 20],
            buttons=[],
        )
        self._fit_dialog(self.auth_dialog, top_pad=16, side_pad=14, bottom_pad=14, spacer_top=0)
        self.auth_dialog.open()

    def _open_create_seller_pin_dialog(self):
        if getattr(self, 'auth_dialog', None):
            try:
                self.auth_dialog.dismiss()
            except Exception:
                pass
            self.auth_dialog = None
        Window.release_all_keyboards()

        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(10),
            size_hint_y=None,
            adaptive_height=True,
            padding=[0, dp(4), 0, 0],
        )
        content.add_widget(
            self._ui_wrap_label(
                'Définissez votre code PIN Administrateur pour protéger le changement de mode.',
                font_style='Caption',
                halign='center',
                theme_color='Secondary',
                min_h=18,
            )
        )

        def _make_pin_box(hint_txt: str):
            card = MDCard(
                orientation='horizontal',
                size_hint=(1, None),
                height=dp(60),
                radius=[12],
                elevation=0,
                md_bg_color=(0.96, 0.97, 0.99, 1),
                line_color=(0.82, 0.88, 0.96, 1),
                line_width=1.1,
                padding=[dp(12), dp(2), dp(6), dp(2)],
                spacing=dp(8),
            )
            card.add_widget(
                MDIcon(
                    icon='shield-key-outline',
                    theme_text_color='Custom',
                    text_color=(0.09, 0.44, 0.84, 1),
                    font_size='20sp',
                    size_hint_x=None,
                    width=dp(22),
                    pos_hint={'center_y': 0.5},
                )
            )
            tf = MDTextField(
                hint_text=hint_txt,
                password=True,
                halign='center',
                font_size='18sp',
                mode='line',
                active_line=False,
                line_color_normal=(0, 0, 0, 0),
                line_color_focus=(0, 0, 0, 0),
                pos_hint={'center_y': 0.5},
            )
            card.add_widget(tf)
            card.add_widget(
                MDIconButton(
                    icon='eye-off-outline',
                    theme_text_color='Secondary',
                    icon_size='20sp',
                    pos_hint={'center_y': 0.5},
                    on_release=lambda b, f=tf: [
                        setattr(f, 'password', not f.password),
                        setattr(b, 'icon', 'eye-off-outline' if f.password else 'eye-outline'),
                    ],
                )
            )
            return card, tf

        c1, pin1_field = _make_pin_box('Nouveau code PIN')
        c2, pin2_field = _make_pin_box('Confirmer le code PIN')
        content.add_widget(c1)
        content.add_widget(c2)

        def _save_new_pin(*args):
            pin1_field.error = False
            pin2_field.error = False
            p1 = (pin1_field.text or '').strip()
            p2 = (pin2_field.text or '').strip()
            if not p1:
                pin1_field.error = True
                self.notify('Veuillez saisir un code PIN.', 'error')
                return
            if p1 != p2:
                pin1_field.error = True
                pin2_field.error = True
                pin2_field.text = ''
                self.notify('Les codes PIN ne correspondent pas.', 'error')
                return

            self.db.set_setting('seller_password', p1)
            pin1_field.focus = False
            pin2_field.focus = False
            Window.release_all_keyboards()
            if getattr(self, 'auth_dialog', None):
                self.auth_dialog.dismiss()
            self.notify('Code PIN créé avec succès.', 'success')
            Clock.schedule_once(lambda dt: self._open_seller_toggle_dialog(), 0.24)

        pin1_field.bind(on_text_validate=lambda x: setattr(pin2_field, 'focus', True))
        pin2_field.bind(on_text_validate=_save_new_pin)

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint=(1, None), height=dp(48), padding=[0, dp(4), 0, 0])
        btns.add_widget(
            self._ui_card_btn(
                'ANNULER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                size_hint_x=0.38,
                font_size='12.5sp',
                on_release=lambda x: self.auth_dialog.dismiss() if getattr(self, 'auth_dialog', None) else None,
            )
        )
        btns.add_widget(
            self._ui_card_btn(
                'ENREGISTRER',
                (0.0, 0.62, 0.25, 1),
                icon='check-circle-outline',
                size_hint_x=0.62,
                font_size='13sp',
                elevation=1,
                on_release=_save_new_pin,
            )
        )
        content.add_widget(btns)

        self.auth_dialog = MDDialog(
            title='Nouveau Code PIN',
            type='custom',
            content_cls=content,
            size_hint=(0.92, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.auth_dialog, top_pad=16, side_pad=14, bottom_pad=14, spacer_top=8)
        self.auth_dialog.open()
        Clock.schedule_once(lambda dt: setattr(pin1_field, 'focus', True), 0.25)

    def _verify_seller_pin(self, stored: str):
        if not getattr(self, 'seller_pass_field', None):
            return
        self.seller_pass_field.error = False
        pwd = (self.seller_pass_field.text or '').strip()
        if not pwd:
            self.seller_pass_field.error = True
            self.notify('Veuillez saisir le code PIN.', 'error')
            return
        if pwd == stored:
            self.seller_pass_field.focus = False
            Window.release_all_keyboards()
            if getattr(self, 'auth_dialog', None):
                self.auth_dialog.dismiss()
            Clock.schedule_once(lambda dt: self._open_seller_toggle_dialog(), 0.24)
        else:
            self.seller_pass_field.error = True
            self.seller_pass_field.text = ''
            self.notify('Code PIN incorrect.', 'error')

    def _open_seller_toggle_dialog(self):
        if getattr(self, '_mode_switch_in_progress', False):
            return
        if getattr(self, 'toggle_dialog', None):
            try:
                self.toggle_dialog.dismiss()
            except Exception:
                pass
            self.toggle_dialog = None
        Window.release_all_keyboards()

        cur_seller = bool(self.is_seller_mode)
        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(10),
            size_hint_y=None,
            adaptive_height=True,
            padding=[0, dp(4), 0, 0],
        )

        def _make_mode_card(title_t: str, sub_t: str, ico_name: str, accent_col, active_bg, is_active: bool, target_val: bool):
            card = MDCard(
                orientation='horizontal',
                size_hint=(1, None),
                height=dp(74),
                radius=[14],
                elevation=0,
                md_bg_color=active_bg if is_active else (0.97, 0.98, 0.99, 1),
                line_color=accent_col if is_active else (0.86, 0.88, 0.93, 1),
                line_width=1.6 if is_active else 1,
                padding=[dp(12), dp(8)],
                spacing=dp(10),
                ripple_behavior=True,
                on_release=lambda x, tv=target_val: self._apply_seller_mode_switch(tv),
            )
            card.add_widget(
                self._ui_icon_pill(
                    ico_name,
                    (accent_col[0], accent_col[1], accent_col[2], 0.14),
                    accent_col,
                    size_dp=42,
                    radius=12,
                    icon_size='22sp',
                )
            )
            txt_box = MDBoxLayout(
                orientation='vertical',
                size_hint_x=1,
                adaptive_height=True,
                spacing=dp(2),
                pos_hint={'center_y': 0.5},
            )
            txt_box.add_widget(
                MDLabel(
                    text=title_t,
                    font_style='Subtitle2',
                    bold=True,
                    theme_text_color='Custom',
                    text_color=accent_col if is_active else (0.14, 0.16, 0.22, 1),
                    adaptive_height=True,
                )
            )
            txt_box.add_widget(
                self._ui_wrap_label(
                    sub_t,
                    font_style='Caption',
                    theme_color='Secondary',
                    min_h=16,
                )
            )
            card.add_widget(txt_box)
            card.add_widget(
                MDIcon(
                    icon='check-circle' if is_active else 'chevron-right',
                    theme_text_color='Custom' if is_active else 'Hint',
                    text_color=accent_col if is_active else (0.55, 0.58, 0.65, 1),
                    font_size='22sp',
                    size_hint=(None, None),
                    size=(dp(24), dp(24)),
                    pos_hint={'center_y': 0.5},
                )
            )
            return card

        content.add_widget(
            _make_mode_card(
                'Mode Administrateur',
                'Accès complet : stock, achats, bénéfices et paramètres.',
                'shield-crown-outline',
                (0.09, 0.44, 0.84, 1),
                (0.91, 0.95, 1, 1),
                is_active=not cur_seller,
                target_val=False,
            )
        )
        content.add_widget(
            _make_mode_card(
                'Mode Vendeur',
                'Accès restreint : ventes, retours clients et encaissements.',
                'account-tie-outline',
                (0.04, 0.58, 0.26, 1),
                (0.91, 0.98, 0.93, 1),
                is_active=cur_seller,
                target_val=True,
            )
        )
        content.add_widget(
            self._ui_card_btn(
                'FERMER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                height=46,
                on_release=lambda x: self.toggle_dialog.dismiss() if getattr(self, 'toggle_dialog', None) else None,
            )
        )

        self.toggle_dialog = MDDialog(
            title='Changer de Mode',
            type='custom',
            content_cls=content,
            size_hint=(0.92, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.toggle_dialog, top_pad=16, side_pad=14, bottom_pad=14, spacer_top=8)
        self.toggle_dialog.open()

    def _apply_seller_mode_switch(self, target_seller: bool):
        target_seller = bool(target_seller)
        if getattr(self, '_mode_switch_in_progress', False):
            return
        Window.release_all_keyboards()
        if getattr(self, 'toggle_dialog', None):
            self.toggle_dialog.dismiss()
        if bool(self.is_seller_mode) == target_seller:
            return
        self._mode_switch_in_progress = True
        Clock.schedule_once(lambda dt: self._commit_seller_mode_switch(target_seller), 0.26)

    def _commit_seller_mode_switch(self, val: bool):
        try:
            self.is_seller_mode = val
            self.db.set_setting('config_seller_mode', str(val))
            self.update_dashboard_layout()
            self.notify(f"Mode Vendeur : {'Activé' if val else 'Désactivé'}", 'info')
        finally:
            Clock.schedule_once(lambda dt: setattr(self, '_mode_switch_in_progress', False), 0.15)

    def _on_seller_mode_checkbox_active(self, instance, val: bool):
        self._apply_seller_mode_switch(bool(val))


    # --- Export / Import Excel ---
    def get_unified_path(self, filename: str) -> str:
        if platform == 'android':
            try:
                from jnius import autoclass
                env = autoclass('android.os.Environment')
                dl = env.getExternalStoragePublicDirectory(env.DIRECTORY_DOWNLOADS).getAbsolutePath()
                os.makedirs(dl, exist_ok=True)
                probe = os.path.join(dl, '.magpro_probe')
                with open(probe, 'w', encoding='utf-8') as f:
                    f.write('1')
                os.remove(probe)
                return os.path.join(dl, filename) if filename else dl
            except Exception:
                try:
                    if PythonActivity is not None:
                        from jnius import autoclass
                        env = autoclass('android.os.Environment')
                        ext_dl = PythonActivity.mActivity.getExternalFilesDir(env.DIRECTORY_DOWNLOADS)
                        if ext_dl is not None:
                            p = ext_dl.getAbsolutePath()
                            os.makedirs(p, exist_ok=True)
                            return os.path.join(p, filename) if filename else p
                except Exception:
                    pass
                return os.path.join(self.user_data_dir, filename) if filename else self.user_data_dir
        dl = os.path.join(os.path.expanduser('~'), 'Downloads')
        os.makedirs(dl, exist_ok=True)
        return os.path.join(dl, filename) if filename else dl

    def _save_bytes_to_downloads(self, filename: str, raw_bytes: bytes, mime_type: str) -> tuple:
        if platform == 'android':
            try:
                from jnius import autoclass
                version_cls = autoclass('android.os.Build$VERSION')
                if version_cls.SDK_INT >= 29 and PythonActivity is not None:
                    MediaStoreDownloads = autoclass('android.provider.MediaStore$Downloads')
                    MediaColumns = autoclass('android.provider.MediaStore$MediaColumns')
                    ContentValues = autoclass('android.content.ContentValues')
                    Environment = autoclass('android.os.Environment')

                    values = ContentValues()
                    values.put(MediaColumns.DISPLAY_NAME, str(filename))
                    values.put(MediaColumns.MIME_TYPE, str(mime_type))
                    values.put(MediaColumns.RELATIVE_PATH, str(Environment.DIRECTORY_DOWNLOADS))

                    resolver = PythonActivity.mActivity.getContentResolver()
                    uri = resolver.insert(MediaStoreDownloads.EXTERNAL_CONTENT_URI, values)
                    if uri is not None:
                        out_stream = resolver.openOutputStream(uri)
                        if out_stream is not None:
                            try:
                                chunk_size = 16384
                                for i in range(0, len(raw_bytes), chunk_size):
                                    out_stream.write(raw_bytes[i:i + chunk_size])
                                out_stream.flush()
                            finally:
                                out_stream.close()
                            return f'Download/{filename}', uri
            except Exception as exc:
                if DEBUG:
                    print(f'[WARN] MediaStore save fallback: {exc}')

        out_dir = self.get_unified_path('')
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, filename)
        with open(out_path, 'wb') as f:
            f.write(raw_bytes)
        return out_path, None

    def _save_file_to_downloads(self, src_path: str, filename: str, mime_type: str) -> tuple:
        if platform == 'android':
            try:
                from jnius import autoclass
                version_cls = autoclass('android.os.Build$VERSION')
                if version_cls.SDK_INT >= 29 and PythonActivity is not None:
                    MediaStoreDownloads = autoclass('android.provider.MediaStore$Downloads')
                    MediaColumns = autoclass('android.provider.MediaStore$MediaColumns')
                    ContentValues = autoclass('android.content.ContentValues')
                    Environment = autoclass('android.os.Environment')

                    values = ContentValues()
                    values.put(MediaColumns.DISPLAY_NAME, str(filename))
                    values.put(MediaColumns.MIME_TYPE, str(mime_type))
                    values.put(MediaColumns.RELATIVE_PATH, str(Environment.DIRECTORY_DOWNLOADS))

                    resolver = PythonActivity.mActivity.getContentResolver()
                    uri = resolver.insert(MediaStoreDownloads.EXTERNAL_CONTENT_URI, values)
                    if uri is not None:
                        out_stream = resolver.openOutputStream(uri)
                        if out_stream is not None:
                            try:
                                with open(src_path, 'rb') as f:
                                    while True:
                                        chunk = f.read(16384)
                                        if not chunk:
                                            break
                                        out_stream.write(chunk)
                                out_stream.flush()
                            finally:
                                out_stream.close()
                            return f'Download/{filename}', uri
            except Exception as exc:
                if DEBUG:
                    print(f'[WARN] MediaStore file save fallback: {exc}')

        out_dir = self.get_unified_path('')
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, filename)
        if os.path.abspath(src_path) != os.path.abspath(out_path):
            shutil.copyfile(src_path, out_path)
        return out_path, None

    def _share_or_open_exported_file(self, saved_path: str, saved_uri, mime_type: str, action: str = 'share'):
        if platform != 'android':
            try:
                if os.path.exists(saved_path):
                    if sys.platform.startswith('win'):
                        os.startfile(saved_path)
                    else:
                        webbrowser.open(saved_path)
            except Exception as exc:
                self.notify(f'Erreur ouverture : {exc}', 'error')
            return
        try:
            from jnius import autoclass, cast
            intent_cls = autoclass('android.content.Intent')
            str_cls = autoclass('java.lang.String')
            target_uri = saved_uri
            if target_uri is None:
                version_cls = autoclass('android.os.Build$VERSION')
                if version_cls.SDK_INT >= 29 and saved_path and os.path.exists(saved_path):
                    _, target_uri = self._save_file_to_downloads(saved_path, os.path.basename(saved_path), mime_type)
                if target_uri is None:
                    strict = autoclass('android.os.StrictMode')
                    strict.setVmPolicy(autoclass('android.os.StrictMode$VmPolicy$Builder')().build())
                    uri_cls = autoclass('android.net.Uri')
                    file_cls = autoclass('java.io.File')
                    target_uri = uri_cls.fromFile(file_cls(saved_path))

            if action == 'view':
                intent = intent_cls(intent_cls.ACTION_VIEW)
                intent.setDataAndType(target_uri, mime_type)
                intent.addFlags(intent_cls.FLAG_GRANT_READ_URI_PERMISSION | intent_cls.FLAG_ACTIVITY_NEW_TASK)
                cast('android.app.Activity', PythonActivity.mActivity).startActivity(intent)
            else:
                intent = intent_cls(intent_cls.ACTION_SEND)
                intent.setType(mime_type)
                intent.putExtra(intent_cls.EXTRA_STREAM, cast('android.os.Parcelable', target_uri))
                intent.addFlags(intent_cls.FLAG_GRANT_READ_URI_PERMISSION)
                try:
                    clip_cls = autoclass('android.content.ClipData')
                    intent.setClipData(clip_cls.newRawUri(str_cls(''), target_uri))
                except Exception:
                    pass
                chooser = intent_cls.createChooser(intent, str_cls(f'Partager : {os.path.basename(saved_path)}'))
                chooser.addFlags(intent_cls.FLAG_GRANT_READ_URI_PERMISSION)
                cast('android.app.Activity', PythonActivity.mActivity).startActivity(chooser)
        except Exception as exc:
            self.notify(f'Erreur partage/ouverture : {exc}', 'error')

    def _list_excel_files(self) -> list:
        files_by_name = {}

        if platform == 'android':
            try:
                from jnius import autoclass
                version_cls = autoclass('android.os.Build$VERSION')
                if version_cls.SDK_INT >= 29 and PythonActivity is not None:
                    MediaStoreDownloads = autoclass('android.provider.MediaStore$Downloads')
                    MediaColumns = autoclass('android.provider.MediaStore$MediaColumns')
                    ContentUris = autoclass('android.content.ContentUris')
                    resolver = PythonActivity.mActivity.getContentResolver()
                    proj = ['_id', MediaColumns.DISPLAY_NAME, MediaColumns.SIZE, MediaColumns.DATE_MODIFIED]
                    cursor = resolver.query(
                        MediaStoreDownloads.EXTERNAL_CONTENT_URI,
                        proj,
                        None,
                        None,
                        f'{MediaColumns.DATE_MODIFIED} DESC',
                    )
                    if cursor is not None:
                        try:
                            idx_id = cursor.getColumnIndex('_id')
                            idx_name = cursor.getColumnIndex(MediaColumns.DISPLAY_NAME)
                            idx_size = cursor.getColumnIndex(MediaColumns.SIZE)
                            idx_date = cursor.getColumnIndex(MediaColumns.DATE_MODIFIED)
                            while cursor.moveToNext():
                                name = str(cursor.getString(idx_name) or '').strip()
                                if not name.lower().endswith('.xlsx') or name.startswith('.') or name.startswith('~$'):
                                    continue
                                row_id = cursor.getLong(idx_id)
                                uri_obj = ContentUris.withAppendedId(MediaStoreDownloads.EXTERNAL_CONTENT_URI, row_id)
                                sz = int(cursor.getLong(idx_size) or 0) if idx_size >= 0 else 0
                                mtime = float(cursor.getLong(idx_date) or 0) if idx_date >= 0 else 0.0
                                files_by_name[name] = {
                                    'name': name,
                                    'path': f'Download/{name}',
                                    'uri': uri_obj,
                                    'mtime': mtime,
                                    'size': sz,
                                }
                        finally:
                            cursor.close()
            except Exception:
                pass

        try:
            dl_dir = self.get_unified_path('')
            if dl_dir and os.path.exists(dl_dir):
                for fn in os.listdir(dl_dir):
                    if not fn.lower().endswith('.xlsx') or fn.startswith('.') or fn.startswith('~$') or fn.startswith('temp_'):
                        continue
                    fp = os.path.join(dl_dir, fn)
                    if not os.path.isfile(fp):
                        continue
                    if fn not in files_by_name:
                        files_by_name[fn] = {
                            'name': fn,
                            'path': fp,
                            'uri': None,
                            'mtime': os.path.getmtime(fp),
                            'size': os.path.getsize(fp),
                        }
        except Exception:
            pass

        items = list(files_by_name.values())
        items.sort(key=lambda it: it['mtime'], reverse=True)
        return items

    def _execute_excel_export(self, on_done=None):
        filename = f"MagPro_Stock_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        temp_xlsx = os.path.join(self.user_data_dir, f'temp_{int(time.time())}.xlsx')
        conn = self.db.get_connection()
        try:
            rows = conn.execute(
                'SELECT product_ref, barcode, name, reference, stock_store, purchase_price, price, price_semi, price_wholesale FROM products WHERE id <> ?',
                (AppConstants.DEFAULT_PRODUCT_ID,),
            ).fetchall()
            if not rows:
                self.notify('Aucun produit à exporter.', 'warning')
                return
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'Produits'
            ws.append(['N° Produit', 'Code-Barres', 'Désignation', 'Référence', 'Stock', 'Prix Achat', 'Prix Détail', 'Prix Demi-Gros', 'Prix Gros'])
            for r in rows:
                ws.append(list(r))
            wb.save(temp_xlsx)
            with open(temp_xlsx, 'rb') as f:
                xlsx_bytes = f.read()
            self._save_bytes_to_downloads(
                filename,
                xlsx_bytes,
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            self.notify(f'Excel enregistré dans Téléchargements : {filename}', 'success')
            if on_done:
                on_done()
        except Exception as exc:
            self.notify(f'Erreur Export : {exc}', 'error')
        finally:
            conn.close()
            if os.path.exists(temp_xlsx):
                try:
                    os.remove(temp_xlsx)
                except Exception:
                    pass

    def perform_export(self, _retry: int = 0):
        if getattr(self, 'settings_menu_dialog', None) and self.settings_menu_dialog.parent:
            self.settings_menu_dialog.dismiss()
            if _retry < 5:
                Clock.schedule_once(lambda dt: self.perform_export(_retry + 1), 0.10)
                return

        if getattr(self, 'dialog', None):
            try:
                self.dialog.dismiss()
            except Exception:
                pass
            self.dialog = None

        excel_mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

        def _fmt_size(sz: int) -> str:
            if sz < 1024:
                return f'{sz} o'
            if sz < 1024 * 1024:
                return f'{sz / 1024:.1f} Ko'
            return f'{sz / (1024 * 1024):.2f} Mo'

        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(10),
            size_hint_y=None,
            height=min(dp(560), Window.height * 0.80),
            padding=[0, dp(2), 0, 0],
        )

        dir_card = MDCard(
            orientation='horizontal',
            size_hint=(1, None),
            height=dp(42),
            radius=[10],
            elevation=0,
            md_bg_color=(0.91, 0.98, 0.93, 1),
            line_color=(0.72, 0.88, 0.76, 1),
            line_width=1,
            padding=[dp(10), 0],
            spacing=dp(8),
        )
        dir_card.add_widget(
            MDIcon(
                icon='folder-outline',
                theme_text_color='Custom',
                text_color=(0.04, 0.58, 0.25, 1),
                font_size='20sp',
                size_hint=(None, None),
                size=(dp(22), dp(22)),
                pos_hint={'center_y': 0.5},
            )
        )
        lbl_count = MDLabel(
            text='Téléchargements (.xlsx)',
            font_style='Caption',
            bold=True,
            theme_text_color='Custom',
            text_color=(0.04, 0.48, 0.20, 1),
            shorten=True,
            valign='middle',
            pos_hint={'center_y': 0.5},
        )
        dir_card.add_widget(lbl_count)
        content.add_widget(dir_card)

        scroll = MDScrollView(size_hint_y=1, bar_width=dp(3))
        lst = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(8), padding=[0, dp(2), dp(2), dp(4)])
        scroll.add_widget(lst)
        content.add_widget(scroll)

        def _populate_xlsx_list():
            lst.clear_widgets()
            xlsx_files = self._list_excel_files()
            lbl_count.text = f'Dossier : Téléchargements ({len(xlsx_files)} fichier(s) .xlsx)'

            if not xlsx_files:
                empty_card = MDCard(
                    orientation='vertical',
                    size_hint=(1, None),
                    height=dp(150),
                    radius=[14],
                    elevation=0,
                    md_bg_color=(0.97, 0.98, 0.99, 1),
                    line_color=(0.88, 0.90, 0.94, 1),
                    line_width=1,
                    padding=dp(16),
                    spacing=dp(8),
                )
                empty_card.add_widget(MDBoxLayout())
                empty_card.add_widget(MDIcon(icon='file-excel', theme_text_color='Hint', font_size='38sp', halign='center', pos_hint={'center_x': 0.5}))
                empty_card.add_widget(MDLabel(text='Aucun fichier .xlsx trouvé dans Téléchargements.', halign='center', font_style='Subtitle2', bold=True, theme_text_color='Secondary', adaptive_height=True))
                empty_card.add_widget(MDBoxLayout())
                lst.add_widget(empty_card)
            else:
                for it in xlsx_files:
                    dt_str = datetime.fromtimestamp(it['mtime']).strftime('%d/%m/%Y %H:%M') if it['mtime'] > 0 else ''
                    sub_txt = f"{dt_str}  •  {_fmt_size(it['size'])}" if dt_str else _fmt_size(it['size'])

                    card = MDCard(
                        orientation='horizontal',
                        size_hint=(1, None),
                        height=dp(66),
                        radius=[12],
                        elevation=0,
                        md_bg_color=(1, 1, 1, 1),
                        line_color=(0.88, 0.90, 0.94, 1),
                        line_width=1,
                        padding=[dp(10), dp(8)],
                        spacing=dp(10),
                        ripple_behavior=True,
                        on_release=lambda x, f=it: self._share_or_open_exported_file(f['path'], f['uri'], excel_mime, action='view'),
                    )
                    card.add_widget(
                        self._ui_icon_pill(
                            'file-excel',
                            (0.91, 0.98, 0.93, 1),
                            (0.06, 0.58, 0.25, 1),
                            size_dp=42,
                            radius=11,
                            icon_size='22sp',
                        )
                    )
                    info = MDBoxLayout(orientation='vertical', size_hint_x=1, adaptive_height=True, spacing=dp(2), pos_hint={'center_y': 0.5})
                    info.add_widget(
                        MDLabel(
                            text=it['name'],
                            font_style='Subtitle2',
                            bold=True,
                            theme_text_color='Primary',
                            shorten=True,
                            shorten_from='center',
                            adaptive_height=True,
                        )
                    )
                    info.add_widget(
                        MDLabel(
                            text=sub_txt,
                            font_style='Caption',
                            theme_text_color='Secondary',
                            adaptive_height=True,
                        )
                    )
                    card.add_widget(info)
                    card.add_widget(
                        self._ui_icon_pill(
                            'share-variant',
                            (0.90, 0.95, 1, 1),
                            (0.09, 0.44, 0.84, 1),
                            on_release=lambda x, f=it: self._share_or_open_exported_file(f['path'], f['uri'], excel_mime, action='share'),
                            size_dp=36,
                            radius=10,
                            icon_size='18sp',
                        )
                    )
                    lst.add_widget(card)

            Clock.schedule_once(lambda dt: setattr(scroll, 'scroll_y', 1.0), 0.05)

        _populate_xlsx_list()

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint=(1, None), height=dp(48))
        btns.add_widget(
            self._ui_card_btn(
                'FERMER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                size_hint_x=0.35,
                font_size='12.5sp',
                on_release=lambda x: self.dialog.dismiss() if getattr(self, 'dialog', None) else None,
            )
        )
        btns.add_widget(
            self._ui_card_btn(
                'EXPORTER (.xlsx)',
                (0.0, 0.62, 0.25, 1),
                icon='file-excel',
                size_hint_x=0.65,
                font_size='13sp',
                elevation=1,
                on_release=lambda x: self._execute_excel_export(on_done=_populate_xlsx_list),
            )
        )
        content.add_widget(btns)

        self.dialog = MDDialog(
            title='Exporter Produits (.xlsx)',
            type='custom',
            content_cls=content,
            size_hint=(0.95, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.dialog, top_pad=16, side_pad=14, bottom_pad=12, spacer_top=8)
        self.dialog.open()

    def import_data_dialog(self):
        if platform == 'android':
            try:
                from android import activity
                from jnius import autoclass, cast

                intent_cls = autoclass('android.content.Intent')
                intent = intent_cls(intent_cls.ACTION_GET_CONTENT)
                intent.setType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                intent.addCategory(intent_cls.CATEGORY_OPENABLE)
                activity.bind(on_activity_result=self._on_android_excel_chosen)
                cast('android.app.Activity', PythonActivity.mActivity).startActivityForResult(intent, 101)
            except Exception as exc:
                self.notify(f'Erreur sélecteur : {exc}', 'error')
            return

        default_path = self.get_unified_path('')
        if not os.path.exists(default_path):
            default_path = self.user_data_dir
        content = MDBoxLayout(orientation='vertical', spacing=10, size_hint_y=None, height=dp(500))
        list_bg = MDCard(orientation='vertical', size_hint_y=1, md_bg_color=(0.15, 0.15, 0.15, 1), radius=[10], padding='5dp')
        self.import_file_chooser = FileChooserListView(path=default_path, filters=['*.xlsx'], size_hint_y=1)
        list_bg.add_widget(self.import_file_chooser)
        content.add_widget(list_bg)
        btns = MDBoxLayout(spacing=10, size_hint_y=None, height=dp(50))
        btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.4, on_release=lambda x: self.import_diag.dismiss()))
        btns.add_widget(self._ui_card_btn('SUIVANT', (0.09, 0.44, 0.84, 1), size_hint_x=0.6, on_release=lambda x: self._pre_process_import(self.import_file_chooser.selection)))
        content.add_widget(btns)
        self.import_diag = MDDialog(title='Importer Produits (.xlsx)', type='custom', content_cls=content, size_hint=(0.95, None), radius=[18, 18, 18, 18])
        self._fit_dialog(self.import_diag)
        self.import_diag.open()

    def _on_android_excel_chosen(self, req_code, res_code, intent):
        from android import activity
        activity.unbind(on_activity_result=self._on_android_excel_chosen)
        if req_code == 101 and res_code == -1 and intent and intent.getData():
            uri = intent.getData()
            threading.Thread(target=self._copy_excel_uri_worker, args=(uri,), daemon=True).start()

    def _copy_excel_uri_worker(self, uri):
        try:
            from jnius import cast
            cur_act = cast('android.app.Activity', PythonActivity.mActivity)
            stream = cur_act.getContentResolver().openInputStream(uri)
            temp_path = os.path.join(self.user_data_dir, 'temp_import.xlsx')
            with open(temp_path, 'wb') as out:
                buf = bytearray(8192)
                while True:
                    n = stream.read(buf)
                    if n == -1:
                        break
                    out.write(buf[:n])
            stream.close()
            Clock.schedule_once(lambda dt: self._pre_process_import([temp_path]), 0)
        except Exception as exc:
            Clock.schedule_once(lambda dt, e=exc: self.notify(f'Erreur lecture Excel : {e}', 'error'), 0)

    def _pre_process_import(self, selection):
        if not selection:
            self.notify('Aucun fichier sélectionné.', 'error')
            return
        filepath = selection[0]
        if getattr(self, 'import_diag', None):
            self.import_diag.dismiss()
        self.notify('Lecture du fichier Excel en cours...', 'info')

        def _read_worker():
            wb = None
            try:
                wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
                rows_iter = wb.active.iter_rows(values_only=True)
                first_row = next(rows_iter, None)
                if not first_row:
                    Clock.schedule_once(lambda dt: self.notify('Fichier Excel vide.', 'error'), 0)
                    return
                headers = [str(h or '').strip() for h in first_row]
                data = [
                    row for row in rows_iter
                    if any(cell is not None and str(cell).strip() != '' for cell in row)
                ]
                if not data:
                    Clock.schedule_once(lambda dt: self.notify('Fichier Excel vide.', 'error'), 0)
                    return
                Clock.schedule_once(lambda dt, h=headers, d=data: self._open_mapping_dialog(h, d), 0)
            except Exception as exc:
                Clock.schedule_once(lambda dt, e=exc: self.notify(f'Erreur lecture : {e}', 'error'), 0)
            finally:
                if wb is not None:
                    try:
                        wb.close()
                    except Exception:
                        pass

        threading.Thread(target=_read_worker, daemon=True).start()

    def _open_mapping_dialog(self, excel_headers: list, excel_data: list):
        disp_opts = ['-- Ignorer --'] + [self.fix_text(h) for h in excel_headers]
        disp_to_real = dict(zip(disp_opts, ['-- Ignorer --'] + excel_headers))
        db_fields = {
            'product_ref': 'N° Produit',
            'name': 'Désignation *',
            'barcode': 'Code-Barres',
            'reference': 'Référence',
            'price': 'Prix Détail',
            'price_semi': 'Prix Demi-Gros',
            'price_wholesale': 'Prix Gros',
            'purchase_price': 'Prix Achat (Coût)',
            'stock_store': 'Stock Initial',
        }
        self.spinners = {}
        self.spinner_map = disp_to_real

        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, height=dp(520))
        content.add_widget(MDLabel(text='Associez les colonnes Excel aux champs :', font_style='Subtitle2', size_hint_y=None, height=dp(30)))
        scroll = MDScrollView(size_hint_y=1)
        grid = MDGridLayout(cols=2, spacing=dp(15), padding=[dp(5), dp(10)], adaptive_height=True)
        for key, label in db_fields.items():
            grid.add_widget(MDLabel(text=label, size_hint_x=0.45, bold=True, size_hint_y=None, height=dp(45)))
            sp = Spinner(text='-- Ignorer --', values=disp_opts, size_hint_x=0.55, size_hint_y=None, height=dp(45), background_normal='', background_color=(0.3, 0.3, 0.3, 1), font_name='ArabicFont')
            for h in excel_headers:
                hl = h.lower()
                match = (
                    (key == 'name' and any(x in hl for x in ('nom', 'name', 'désignation', 'libelle')))
                    or (key == 'product_ref' and any(x in hl for x in ('n° produit', 'numéro', 'n° article')))
                    or (key == 'barcode' and any(x in hl for x in ('code', 'bar', 'ean')))
                    or (key == 'reference' and 'réf' in hl)
                    or (key == 'stock_store' and any(x in hl for x in ('stock', 'qté', 'quantité')))
                    or (key == 'purchase_price' and any(x in hl for x in ('achat', 'cost', 'coût')))
                    or (key == 'price' and ('détail' in hl or hl == 'prix'))
                    or (key == 'price_semi' and 'demi' in hl)
                    or (key == 'price_wholesale' and 'gros' in hl and 'demi' not in hl)
                )
                if match:
                    sp.text = self.fix_text(h)
                    sp.background_color = (0, 0.5, 0.2, 1)
                    break
            self.spinners[key] = sp
            grid.add_widget(sp)
        scroll.add_widget(grid)
        content.add_widget(scroll)

        btns = MDBoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
        btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.4, on_release=lambda x: self.mapping_diag.dismiss()))
        btns.add_widget(self._ui_card_btn('IMPORTER', (0, 0.6, 0.2, 1), icon='file-import', size_hint_x=0.6, elevation=1, on_release=lambda x: self._finalize_import(excel_headers, excel_data)))
        content.add_widget(btns)
        self.mapping_diag = MDDialog(title='Mappage Excel', type='custom', content_cls=content, size_hint=(0.95, None), radius=[18, 18, 18, 18])
        self._fit_dialog(self.mapping_diag)
        self.mapping_diag.open()

    def _finalize_import(self, excel_headers: list, excel_data: list):
        if getattr(self, '_import_in_progress', False):
            return
        indices = {}
        for field, sp in self.spinners.items():
            if sp.text != '-- Ignorer --':
                real = self.spinner_map.get(sp.text)
                if real in excel_headers:
                    indices[field] = excel_headers.index(real)
        if 'name' not in indices:
            self.notify("Le champ 'Désignation' est obligatoire.", 'error')
            return

        if getattr(self, 'mapping_diag', None):
            self.mapping_diag.dismiss()

        self._import_in_progress = True
        self._import_cancel_requested = False
        total_rows = len(excel_data)
        self._show_import_progress_dialog(total_rows)

        user_name = self.current_user_name
        threading.Thread(
            target=self._import_excel_worker,
            args=(excel_data, indices, user_name),
            daemon=True,
        ).start()

    def _show_import_progress_dialog(self, total_rows: int):
        from kivymd.uix.progressbar import MDProgressBar

        if getattr(self, 'import_progress_dialog', None):
            try:
                self.import_progress_dialog.dismiss()
            except Exception:
                pass

        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(12),
            size_hint_y=None,
            adaptive_height=True,
            padding=[0, dp(4), 0, dp(2)],
        )

        badge = self._ui_icon_pill(
            'file-import',
            (0.90, 0.98, 0.92, 1),
            (0.0, 0.62, 0.25, 1),
            size_dp=56,
            radius=28,
            icon_size='28sp',
        )
        badge.pos_hint = {'center_x': 0.5}
        content.add_widget(badge)

        content.add_widget(
            MDLabel(
                text='Importation en cours...',
                halign='center',
                font_style='H6',
                bold=True,
                theme_text_color='Primary',
                size_hint_y=None,
                height=dp(28),
            )
        )

        lbl_item = MDLabel(
            text='Préparation...',
            halign='center',
            font_style='Caption',
            bold=True,
            theme_text_color='Secondary',
            font_name='ArabicFont',
            shorten=True,
            shorten_from='right',
            max_lines=1,
            size_hint_y=None,
            height=dp(20),
        )
        content.add_widget(lbl_item)

        bar_card = MDCard(
            orientation='vertical',
            size_hint=(1, None),
            adaptive_height=True,
            radius=[12],
            elevation=0,
            md_bg_color=(0.96, 0.97, 0.99, 1),
            line_color=(0.86, 0.89, 0.94, 1),
            line_width=1,
            padding=[dp(12), dp(10)],
            spacing=dp(8),
        )

        progress_bar = MDProgressBar(
            max=100,
            value=0,
            size_hint_y=None,
            height=dp(8),
            color=(0.0, 0.62, 0.25, 1),
            back_color=(0.86, 0.89, 0.94, 1),
        )
        bar_card.add_widget(progress_bar)

        pct_row = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(20))
        lbl_pct = MDLabel(
            text='0%',
            halign='left',
            font_style='Subtitle2',
            bold=True,
            theme_text_color='Custom',
            text_color=(0.0, 0.58, 0.22, 1),
        )
        lbl_cnt = MDLabel(
            text=f'0 / {total_rows}',
            halign='right',
            font_style='Caption',
            bold=True,
            theme_text_color='Secondary',
        )
        pct_row.add_widget(lbl_pct)
        pct_row.add_widget(lbl_cnt)
        bar_card.add_widget(pct_row)
        content.add_widget(bar_card)

        stats_row = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint=(1, None), height=dp(52))
        c_ok = MDCard(
            orientation='vertical',
            size_hint_x=0.5,
            radius=[12],
            elevation=0,
            md_bg_color=(0.91, 0.98, 0.93, 1),
            line_color=(0.72, 0.88, 0.76, 1),
            line_width=1,
            padding=[dp(8), dp(6)],
        )
        c_ok.add_widget(MDLabel(text='AJOUTÉS', halign='center', font_style='Caption', bold=True, theme_text_color='Custom', text_color=(0.04, 0.52, 0.20, 1), size_hint_y=None, height=dp(16)))
        lbl_added = MDLabel(text='0', halign='center', valign='middle', font_style='Subtitle1', bold=True, theme_text_color='Custom', text_color=(0.04, 0.58, 0.22, 1))
        c_ok.add_widget(lbl_added)

        c_skip = MDCard(
            orientation='vertical',
            size_hint_x=0.5,
            radius=[12],
            elevation=0,
            md_bg_color=(1, 0.94, 0.94, 1),
            line_color=(0.94, 0.78, 0.78, 1),
            line_width=1,
            padding=[dp(8), dp(6)],
        )
        c_skip.add_widget(MDLabel(text='IGNORÉS', halign='center', font_style='Caption', bold=True, theme_text_color='Custom', text_color=(0.75, 0.14, 0.14, 1), size_hint_y=None, height=dp(16)))
        lbl_skipped = MDLabel(text='0', halign='center', valign='middle', font_style='Subtitle1', bold=True, theme_text_color='Custom', text_color=(0.82, 0.12, 0.12, 1))
        c_skip.add_widget(lbl_skipped)

        stats_row.add_widget(c_ok)
        stats_row.add_widget(c_skip)
        content.add_widget(stats_row)

        def _on_cancel_tap(*args):
            if self._import_cancel_requested:
                return
            self._import_cancel_requested = True
            btn_cancel.disabled = True
            btn_cancel.md_bg_color = (0.85, 0.86, 0.89, 1)
            btn_cancel._lbl.text = 'ARRÊT EN COURS...'

        btn_cancel = self._ui_card_btn(
            'ANNULER',
            (0.92, 0.93, 0.95, 1),
            text_color=(0.82, 0.12, 0.12, 1),
            icon='stop-circle-outline',
            height=46,
            on_release=_on_cancel_tap,
        )
        content.add_widget(btn_cancel)

        self._import_progress_widgets = {
            'bar': progress_bar,
            'lbl_pct': lbl_pct,
            'lbl_cnt': lbl_cnt,
            'lbl_item': lbl_item,
            'lbl_added': lbl_added,
            'lbl_skipped': lbl_skipped,
            'btn_cancel': btn_cancel,
        }

        self.import_progress_dialog = MDDialog(
            title='',
            type='custom',
            content_cls=content,
            size_hint=(0.90, None),
            auto_dismiss=False,
            radius=[20, 20, 20, 20],
            buttons=[],
        )
        self.import_progress_dialog.open()


    def _update_import_progress_ui(self, done: int, total: int, added: int, skipped: int, item_name: str):
        widgets = getattr(self, '_import_progress_widgets', None)
        if not widgets or not getattr(self, 'import_progress_dialog', None):
            return
        pct = int((done / max(1, total)) * 100)
        widgets['bar'].value = pct
        widgets['lbl_pct'].text = f'{pct}%'
        widgets['lbl_cnt'].text = f'{done} / {total}'
        widgets['lbl_added'].text = str(added)
        widgets['lbl_skipped'].text = str(skipped)
        if item_name:
            widgets['lbl_item'].text = self.fix_text(item_name)

    def _import_excel_worker(self, excel_data: list, indices: dict, user_name: str):
        t_start = time.monotonic()
        total_rows = len(excel_data)
        added = 0
        skipped = 0
        errors = []
        cancelled = False
        last_ui_ts = 0.0

        def _clean_str(val) -> str:
            if val is None:
                return ''
            if isinstance(val, float):
                if math.isnan(val) or math.isinf(val):
                    return ''
                if val.is_integer():
                    return str(int(val))
            return str(val).strip()

        try:
            ref_seed = self.db.get_next_product_ref()
            next_ref_num = int(ref_seed) if str(ref_seed).isdigit() else 1
        except Exception:
            next_ref_num = 1

        try:
            for idx, row in enumerate(excel_data, start=1):
                if getattr(self, '_import_cancel_requested', False):
                    cancelled = True
                    break

                def _val(f, default=''):
                    col_idx = indices.get(f)
                    if col_idx is None or col_idx >= len(row) or row[col_idx] is None:
                        return default
                    return row[col_idx]

                name_v = _clean_str(_val('name', ''))
                if not name_v:
                    skipped += 1
                    errors.append(f'Ligne {idx + 1} : Désignation manquante.')
                    now_ts = time.monotonic()
                    if (now_ts - last_ui_ts) >= 0.06 or idx == total_rows:
                        last_ui_ts = now_ts
                        Clock.schedule_once(
                            lambda dt, d=idx, t=total_rows, a=added, s=skipped: self._update_import_progress_ui(d, t, a, s, 'Ligne ignorée'),
                            0,
                        )
                    continue

                raw_ref = _clean_str(_val('product_ref', ''))
                used_auto_ref = False
                if not raw_ref:
                    raw_ref = f'{next_ref_num:05d}'
                    used_auto_ref = True

                try:
                    self.db.save_product(
                        {
                            'action': 'add',
                            'name': name_v,
                            'barcode': _clean_str(_val('barcode', '')),
                            'product_ref': raw_ref,
                            'reference': _clean_str(_val('reference', '')),
                            'price': FinancialEngine.to_decimal(_val('price', 0)),
                            'price_semi': FinancialEngine.to_decimal(_val('price_semi', 0)),
                            'price_wholesale': FinancialEngine.to_decimal(_val('price_wholesale', 0)),
                            'purchase_price': FinancialEngine.to_decimal(_val('purchase_price', 0)),
                            'stock_store': FinancialEngine.to_decimal(_val('stock_store', 0)),
                            'is_unlimited': False,
                        },
                        user_name=user_name,
                    )
                    added += 1
                    if used_auto_ref:
                        next_ref_num += 1
                    elif raw_ref.isdigit():
                        next_ref_num = max(next_ref_num, int(raw_ref) + 1)
                except Exception as exc:
                    skipped += 1
                    errors.append(f'Ligne {idx + 1} ({name_v}) : {exc}')

                now_ts = time.monotonic()
                if (now_ts - last_ui_ts) >= 0.06 or idx == total_rows:
                    last_ui_ts = now_ts
                    Clock.schedule_once(
                        lambda dt, d=idx, t=total_rows, a=added, s=skipped, nm=name_v: self._update_import_progress_ui(d, t, a, s, nm),
                        0,
                    )

                if idx % 5 == 0:
                    time.sleep(0.001)
        except Exception as fatal_exc:
            errors.append(f'Erreur inattendue : {fatal_exc}')

        elapsed = time.monotonic() - t_start
        Clock.schedule_once(
            lambda dt: self._show_import_report_dialog(total_rows, added, skipped, errors, elapsed, cancelled),
            0,
        )


    def _show_import_report_dialog(self, total_rows: int, added: int, skipped: int, errors: list, elapsed_sec: float, cancelled: bool = False):
        self._import_in_progress = False
        self._import_cancel_requested = False

        if getattr(self, 'import_progress_dialog', None):
            try:
                self.import_progress_dialog.dismiss()
            except Exception:
                pass
            self.import_progress_dialog = None
        self._import_progress_widgets = None

        self.check_and_load_stats()
        if getattr(self, 'rv_products', None) and getattr(self, 'sm', None) and self.sm.current == 'products':
            self.load_more_products(reset=True)

        if cancelled:
            title_txt = 'Importation interrompue'
            accent = (0.88, 0.45, 0.0, 1)
            soft_bg = (1, 0.96, 0.90, 1)
            icon_name = 'alert-circle-outline'
            self.notify(f'Import annulé : {added} ajoutés ({skipped} ignorés).', 'warning')
        elif added > 0 and skipped == 0:
            title_txt = 'Importation réussie !'
            accent = (0.04, 0.62, 0.25, 1)
            soft_bg = (0.90, 0.98, 0.92, 1)
            icon_name = 'check-decagram'
            self.notify(f'Import terminé : {added} produit(s) ajouté(s).', 'success')
        elif added > 0:
            title_txt = 'Importation terminée'
            accent = (0.09, 0.44, 0.84, 1)
            soft_bg = (0.90, 0.95, 1, 1)
            icon_name = 'file-check-outline'
            self.notify(f'Import terminé : {added} ajoutés ({skipped} ignorés).', 'warning')
        else:
            title_txt = 'Aucun produit importé'
            accent = (0.82, 0.12, 0.12, 1)
            soft_bg = (1, 0.93, 0.93, 1)
            icon_name = 'close-circle-outline'
            self.notify(f'Échec import : {skipped} ligne(s) ignorée(s).', 'error')

        if getattr(self, 'import_report_dialog', None):
            try:
                self.import_report_dialog.dismiss()
            except Exception:
                pass

        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(11),
            size_hint_y=None,
            adaptive_height=True,
            padding=[0, dp(2), 0, dp(2)],
        )

        badge = self._ui_icon_pill(icon_name, soft_bg, accent, size_dp=56, radius=28, icon_size='28sp')
        badge.pos_hint = {'center_x': 0.5}
        content.add_widget(badge)

        content.add_widget(
            MDLabel(
                text=title_txt,
                halign='center',
                font_style='H6',
                bold=True,
                theme_text_color='Custom',
                text_color=accent,
                size_hint_y=None,
                height=dp(28),
            )
        )

        content.add_widget(
            MDLabel(
                text=f'Durée : {elapsed_sec:.1f} s',
                halign='center',
                font_style='Caption',
                bold=True,
                theme_text_color='Secondary',
                size_hint_y=None,
                height=dp(18),
            )
        )

        kpi_row = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint=(1, None), height=dp(62))

        def _kpi_box(lbl_t: str, val_t: str, bg_c, bd_c, tx_c):
            c = MDCard(
                orientation='vertical',
                size_hint_x=1,
                radius=[12],
                elevation=0,
                md_bg_color=bg_c,
                line_color=bd_c,
                line_width=1,
                padding=[dp(6), dp(6)],
            )
            c.add_widget(MDLabel(text=lbl_t, halign='center', font_style='Caption', bold=True, theme_text_color='Secondary', size_hint_y=None, height=dp(16)))
            c.add_widget(MDLabel(text=val_t, halign='center', valign='middle', font_style='H6', font_size='18sp', bold=True, theme_text_color='Custom', text_color=tx_c))
            return c

        kpi_row.add_widget(_kpi_box('LIGNES', str(total_rows), (0.95, 0.96, 0.98, 1), (0.86, 0.88, 0.93, 1), (0.12, 0.14, 0.18, 1)))
        kpi_row.add_widget(_kpi_box('AJOUTÉS', str(added), (0.91, 0.98, 0.93, 1), (0.72, 0.88, 0.76, 1), (0.04, 0.58, 0.22, 1)))
        kpi_row.add_widget(_kpi_box('IGNORÉS', str(skipped), (1, 0.94, 0.94, 1), (0.94, 0.78, 0.78, 1), (0.82, 0.12, 0.12, 1)))
        content.add_widget(kpi_row)

        if errors:
            err_card = MDCard(
                orientation='vertical',
                size_hint=(1, None),
                height=min(dp(200), Window.height * 0.30),
                radius=[12],
                elevation=0,
                md_bg_color=(0.98, 0.98, 0.99, 1),
                line_color=(0.90, 0.84, 0.84, 1),
                line_width=1,
                padding=[dp(10), dp(8)],
                spacing=dp(6),
            )
            err_card.add_widget(
                MDLabel(
                    text=f'Détails des lignes ignorées ({len(errors)}) :',
                    font_style='Caption',
                    bold=True,
                    theme_text_color='Custom',
                    text_color=(0.75, 0.14, 0.14, 1),
                    size_hint_y=None,
                    height=dp(18),
                )
            )
            scr = MDScrollView(size_hint_y=1, bar_width=dp(3))
            err_list_box = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(4), padding=[dp(2), dp(2), dp(4), dp(2)])
            max_display = 100
            for msg in errors[:max_display]:
                err_list_box.add_widget(
                    self._ui_wrap_label(
                        f'• {self.fix_text(msg)}',
                        font_style='Caption',
                        theme_color='Secondary',
                        min_h=16,
                    )
                )
                err_list_box.add_widget(MDSeparator(height=dp(1), color=(0.92, 0.93, 0.95, 1)))
            if len(errors) > max_display:
                err_list_box.add_widget(
                    MDLabel(
                        text=f'... et {len(errors) - max_display} autre(s) ligne(s).',
                        font_style='Caption',
                        bold=True,
                        theme_text_color='Hint',
                        size_hint_y=None,
                        height=dp(18),
                    )
                )
            scr.add_widget(err_list_box)
            err_card.add_widget(scr)
            content.add_widget(err_card)

            report_txt = (
                f'=== Rapport Import Excel ===\n'
                f'Statut : {title_txt}\n'
                f'Lignes traitées : {total_rows}\n'
                f'Ajoutés : {added}\n'
                f'Ignorés : {skipped}\n'
                f'Durée : {elapsed_sec:.1f} s\n\n'
                + '\n'.join(errors)
            )
            btns = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint=(1, None), height=dp(48))
            btns.add_widget(
                self._ui_card_btn(
                    'COPIER RAPPORT',
                    (0.90, 0.95, 1, 1),
                    text_color=(0.09, 0.44, 0.84, 1),
                    icon='content-copy',
                    size_hint_x=0.58,
                    font_size='12sp',
                    on_release=lambda x: [Clipboard.copy(report_txt), self.notify('Rapport copié.', 'success')],
                )
            )
            btns.add_widget(
                self._ui_card_btn(
                    'FERMER',
                    (0.09, 0.44, 0.84, 1),
                    icon='check-circle-outline',
                    size_hint_x=0.42,
                    font_size='12.5sp',
                    elevation=1,
                    on_release=lambda x: self.import_report_dialog.dismiss() if getattr(self, 'import_report_dialog', None) else None,
                )
            )
            content.add_widget(btns)
        else:
            content.add_widget(
                self._ui_card_btn(
                    'FERMER',
                    (0.04, 0.62, 0.25, 1),
                    icon='check-circle-outline',
                    elevation=1,
                    on_release=lambda x: self.import_report_dialog.dismiss() if getattr(self, 'import_report_dialog', None) else None,
                )
            )

        self.import_report_dialog = MDDialog(
            title='',
            type='custom',
            content_cls=content,
            size_hint=(0.94, None),
            radius=[20, 20, 20, 20],
            buttons=[],
        )
        self._fit_dialog(self.import_report_dialog, top_pad=16, side_pad=14, bottom_pad=14, spacer_top=0)
        self.import_report_dialog.open()

    # --- Sauvegarde & Restauration ZIP (DB + Images + Logo) ---
    def get_backup_directory(self) -> str:
        try:
            bdir = os.path.join(self.get_unified_path(''), 'MagPro_Backups')
            os.makedirs(bdir, exist_ok=True)
            test_f = os.path.join(bdir, '.write_test')
            with open(test_f, 'w', encoding='utf-8') as f:
                f.write('1')
            os.remove(test_f)
            return bdir
        except Exception:
            fallback = os.path.join(self.user_data_dir, 'MagPro_Backups')
            os.makedirs(fallback, exist_ok=True)
            return fallback

    def _rotate_backups(self, bdir: str):
        try:
            files = os.listdir(bdir)
            for prefix, max_keep in (('Backup_Daily_', 30), ('Backup_Rolling_', 10)):
                matched = sorted([os.path.join(bdir, f) for f in files if f.startswith(prefix) and f.endswith('.zip')], key=os.path.getmtime)
                while len(matched) > max_keep:
                    os.remove(matched.pop(0))
        except Exception:
            pass

    def _write_backup_zip(self, zip_path: str, progress_cb=None):
        temp_db = os.path.join(self.user_data_dir, f'temp_bk_{int(time.time())}_{random.randint(100, 999)}.db')
        stage_zip = f'{zip_path}.tmp'
        try:
            if progress_cb:
                progress_cb(5, 'Copie base de données...', force=True)

            src_conn = sqlite3.connect(self.db.db_path)
            dst_conn = sqlite3.connect(temp_db)

            def _sqlite_prog(status, remaining, total):
                if progress_cb and total > 0:
                    frac = max(0.0, min(1.0, (total - remaining) / float(total)))
                    progress_cb(5 + int(25 * frac), 'Copie base de données...')

            src_conn.backup(dst_conn, pages=50, progress=_sqlite_prog if progress_cb else None)
            dst_conn.close()
            src_conn.close()

            media_files = []
            for folder in ('product_images', 'store_assets'):
                f_dir = os.path.join(self.user_data_dir, folder)
                if os.path.exists(f_dir):
                    for root, _, files in os.walk(f_dir):
                        for f in files:
                            full_p = os.path.join(root, f)
                            rel_p = os.path.relpath(full_p, self.user_data_dir).replace('\\', '/')
                            try:
                                sz = os.path.getsize(full_p)
                            except Exception:
                                sz = 0
                            media_files.append((full_p, rel_p, sz))

            db_sz = os.path.getsize(temp_db) if os.path.exists(temp_db) else 0
            total_bytes = max(1, db_sz + sum(sz for _, _, sz in media_files))
            done_bytes = 0

            if progress_cb:
                progress_cb(32, 'Compression base de données...', force=True)

            with zipfile.ZipFile(stage_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                zf.write(temp_db, arcname=AppConstants.DB_NAME)
                done_bytes += db_sz
                if progress_cb:
                    progress_cb(32 + int(66 * min(1.0, done_bytes / total_bytes)), 'Compression base de données...', force=True)

                total_media = len(media_files)
                for idx, (full_p, arc_n, sz) in enumerate(media_files, start=1):
                    if os.path.exists(full_p):
                        zf.write(full_p, arcname=arc_n)
                    done_bytes += sz
                    if progress_cb:
                        progress_cb(
                            32 + int(66 * min(1.0, done_bytes / total_bytes)),
                            f'Compression médias ({idx}/{total_media})...',
                            force=(idx == total_media),
                        )

            os.replace(stage_zip, zip_path)
            if progress_cb:
                progress_cb(100, 'Terminé !', force=True)
        except Exception:
            if os.path.exists(stage_zip):
                try:
                    os.remove(stage_zip)
                except Exception:
                    pass
            raise
        finally:
            if os.path.exists(temp_db):
                try:
                    os.remove(temp_db)
                except Exception:
                    pass

    def perform_local_backup(self, backup_type: str = 'manual'):
        if backup_type == 'manual':
            if getattr(self, '_backup_in_progress', False):
                return
            self._backup_in_progress = True
            now = datetime.now()
            fname = f"MagPro_Backup_{now.strftime('%Y-%m-%d_%H-%M-%S')}.zip"
            self._show_backup_progress_dialog(fname)
            threading.Thread(target=self._backup_worker, args=(fname,), daemon=True).start()
            return

        try:
            bdir = self.get_backup_directory()
            now = datetime.now()
            if backup_type == 'daily':
                fname = f"Backup_Daily_{now.strftime('%Y-%m-%d')}.zip"
            elif backup_type == 'rolling':
                fname = f"Backup_Rolling_{now.strftime('%Y-%m-%d_%H-%M-%S')}.zip"
            else:
                fname = f"MagPro_Backup_{now.strftime('%Y-%m-%d_%H-%M-%S')}.zip"

            zip_path = os.path.join(bdir, fname)
            self._write_backup_zip(zip_path)
            self._rotate_backups(bdir)
        except Exception:
            pass

    def _show_backup_progress_dialog(self, filename: str):
        from kivymd.uix.progressbar import MDProgressBar

        if getattr(self, 'backup_progress_dialog', None):
            try:
                self.backup_progress_dialog.dismiss()
            except Exception:
                pass

        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(12),
            size_hint_y=None,
            adaptive_height=True,
            padding=[0, dp(4), 0, dp(4)],
        )
        badge = self._ui_icon_pill(
            'database-export',
            (0.94, 0.91, 1, 1),
            (0.45, 0.2, 0.72, 1),
            size_dp=56,
            radius=28,
            icon_size='28sp',
        )
        badge.pos_hint = {'center_x': 0.5}
        content.add_widget(badge)

        content.add_widget(
            MDLabel(
                text='Sauvegarde en cours...',
                halign='center',
                font_style='H6',
                bold=True,
                theme_text_color='Primary',
                size_hint_y=None,
                height=dp(28),
            )
        )
        content.add_widget(
            MDLabel(
                text=filename,
                halign='center',
                font_style='Caption',
                bold=True,
                theme_text_color='Secondary',
                shorten=True,
                shorten_from='center',
                size_hint_y=None,
                height=dp(18),
            )
        )

        bar_card = MDCard(
            orientation='vertical',
            size_hint=(1, None),
            adaptive_height=True,
            radius=[12],
            elevation=0,
            md_bg_color=(0.96, 0.97, 0.99, 1),
            line_color=(0.86, 0.89, 0.94, 1),
            line_width=1,
            padding=[dp(12), dp(12)],
            spacing=dp(8),
        )
        progress_bar = MDProgressBar(
            max=100,
            value=0,
            size_hint_y=None,
            height=dp(8),
            color=(0.45, 0.2, 0.72, 1),
            back_color=(0.86, 0.89, 0.94, 1),
        )
        bar_card.add_widget(progress_bar)

        pct_row = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(20))
        lbl_pct = MDLabel(
            text='0%',
            halign='left',
            font_style='Subtitle2',
            bold=True,
            theme_text_color='Custom',
            text_color=(0.45, 0.2, 0.72, 1),
            size_hint_x=0.30,
        )
        lbl_step = MDLabel(
            text='Préparation...',
            halign='right',
            font_style='Caption',
            bold=True,
            theme_text_color='Secondary',
            shorten=True,
            size_hint_x=0.70,
        )
        pct_row.add_widget(lbl_pct)
        pct_row.add_widget(lbl_step)
        bar_card.add_widget(pct_row)
        content.add_widget(bar_card)

        self._backup_progress_widgets = {
            'bar': progress_bar,
            'lbl_pct': lbl_pct,
            'lbl_step': lbl_step,
        }

        self.backup_progress_dialog = MDDialog(
            title='',
            type='custom',
            content_cls=content,
            size_hint=(0.90, None),
            auto_dismiss=False,
            radius=[20, 20, 20, 20],
            buttons=[],
        )
        self._fit_dialog(self.backup_progress_dialog, top_pad=16, side_pad=14, bottom_pad=14, spacer_top=0)
        self.backup_progress_dialog.open()

    def _update_backup_progress_ui(self, pct: int, step_text: str):
        widgets = getattr(self, '_backup_progress_widgets', None)
        if not widgets or not getattr(self, 'backup_progress_dialog', None):
            return
        pct = max(0, min(100, int(pct)))
        widgets['bar'].value = pct
        widgets['lbl_pct'].text = f'{pct}%'
        widgets['lbl_step'].text = step_text

    def _backup_worker(self, fname: str):
        last_ui_ts = 0.0

        def _push_ui(pct: int, text: str, force: bool = False):
            nonlocal last_ui_ts
            now_ts = time.monotonic()
            if force or (now_ts - last_ui_ts) >= 0.05:
                last_ui_ts = now_ts
                Clock.schedule_once(lambda dt, p=pct, t=text: self._update_backup_progress_ui(p, t), 0)

        try:
            bdir = self.get_backup_directory()
            zip_path = os.path.join(bdir, fname)
            self._write_backup_zip(zip_path, progress_cb=_push_ui)
            self._rotate_backups(bdir)
            Clock.schedule_once(lambda dt: self._finish_backup_ui(fname, None), 0.12)
        except Exception as exc:
            Clock.schedule_once(lambda dt, e=str(exc): self._finish_backup_ui(None, e), 0)

    def _finish_backup_ui(self, fname: str | None, error_msg: str | None):
        if getattr(self, 'backup_progress_dialog', None):
            try:
                self.backup_progress_dialog.dismiss()
            except Exception:
                pass
            self.backup_progress_dialog = None
        self._backup_progress_widgets = None
        self._backup_in_progress = False

        if error_msg or not fname:
            self.notify(f'Erreur sauvegarde : {error_msg or "échec"}', 'error')
        else:
            self.notify(f'Sauvegarde enregistrée : {fname}', 'success')

    def trigger_auto_backups(self, dt=None):
        self.perform_local_backup('daily')
        self.perform_local_backup('rolling')

    def share_database_file(self):
        if getattr(self, '_share_in_progress', False):
            return
        self._share_in_progress = True
        fname = f"MagPro_Share_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
        tmp_path = self._get_temp_share_path(fname)
        self.notify('Préparation du partage...', 'info')

        def _share_worker():
            try:
                self._write_backup_zip(tmp_path)
                Clock.schedule_once(lambda dt: self._finish_share_db(tmp_path, None), 0)
            except Exception as exc:
                Clock.schedule_once(lambda dt, e=str(exc): self._finish_share_db(None, e), 0)

        threading.Thread(target=_share_worker, daemon=True).start()

    def _finish_share_db(self, tmp_path: str | None, error_msg: str | None):
        self._share_in_progress = False
        if error_msg or not tmp_path:
            self.notify(f'Erreur partage : {error_msg or "échec"}', 'error')
            return
        if platform == 'android':
            try:
                fname = os.path.basename(tmp_path)
                saved_path, saved_uri = self._save_file_to_downloads(tmp_path, fname, 'application/zip')
                if os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except Exception:
                        pass
                self._share_or_open_exported_file(saved_path, saved_uri, 'application/zip', action='share')
            except Exception as exc:
                self.notify(f'Erreur partage : {exc}', 'error')
        else:
            try:
                if sys.platform.startswith('win'):
                    import subprocess
                    subprocess.Popen(['explorer', '/select,', os.path.normpath(tmp_path)])
                else:
                    webbrowser.open(os.path.dirname(tmp_path))
                self.notify('Archive prête au partage (temporaire, sans sauvegarde).', 'success')
            except Exception:
                self._share_or_open_exported_file(tmp_path, None, 'application/zip', action='share')

    def show_restore_dialog(self):
        if getattr(self, '_restore_in_progress', False):
            return
        if getattr(self, 'restore_dialog', None):
            try:
                self.restore_dialog.dismiss()
            except Exception:
                pass

        bdir = self.get_backup_directory()
        search_dirs = [bdir]
        dl_bdir = os.path.join(self.get_unified_path(''), 'MagPro_Backups')
        if os.path.abspath(dl_bdir) != os.path.abspath(bdir) and os.path.exists(dl_bdir):
            search_dirs.append(dl_bdir)

        backups = []
        seen_paths = set()
        for d in search_dirs:
            if not os.path.exists(d):
                continue
            try:
                for fn in os.listdir(d):
                    if fn.startswith('.') or fn.startswith('temp_'):
                        continue
                    if not fn.lower().endswith(('.zip', '.db')):
                        continue
                    fp = os.path.join(d, fn)
                    ap = os.path.abspath(fp)
                    if ap in seen_paths or not os.path.isfile(fp):
                        continue
                    seen_paths.add(ap)
                    backups.append((fp, fn, os.path.getmtime(fp), os.path.getsize(fp)))
            except Exception:
                pass
        backups.sort(key=lambda item: item[2], reverse=True)

        def _fmt_size(sz: int) -> str:
            if sz < 1024:
                return f'{sz} o'
            if sz < 1024 * 1024:
                return f'{sz / 1024:.1f} Ko'
            return f'{sz / (1024 * 1024):.2f} Mo'

        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(10),
            size_hint_y=None,
            height=min(dp(560), Window.height * 0.80),
            padding=[0, dp(2), 0, 0],
        )

        dir_card = MDCard(
            orientation='horizontal',
            size_hint=(1, None),
            height=dp(42),
            radius=[10],
            elevation=0,
            md_bg_color=(0.94, 0.97, 1, 1),
            line_color=(0.80, 0.88, 0.97, 1),
            line_width=1,
            padding=[dp(10), 0],
            spacing=dp(8),
        )
        dir_card.add_widget(
            MDIcon(
                icon='folder-clock-outline',
                theme_text_color='Custom',
                text_color=(0.09, 0.44, 0.84, 1),
                font_size='20sp',
                size_hint=(None, None),
                size=(dp(22), dp(22)),
                pos_hint={'center_y': 0.5},
            )
        )
        dir_card.add_widget(
            MDLabel(
                text=f'Dossier : MagPro_Backups ({len(backups)})',
                font_style='Caption',
                bold=True,
                theme_text_color='Custom',
                text_color=(0.08, 0.35, 0.72, 1),
                shorten=True,
                valign='middle',
                pos_hint={'center_y': 0.5},
            )
        )
        content.add_widget(dir_card)

        scroll = MDScrollView(size_hint_y=1, bar_width=dp(3))
        lst = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(8), padding=[0, dp(2), dp(2), dp(4)])

        if not backups:
            empty_card = MDCard(
                orientation='vertical',
                size_hint=(1, None),
                height=dp(150),
                radius=[14],
                elevation=0,
                md_bg_color=(0.97, 0.98, 0.99, 1),
                line_color=(0.88, 0.90, 0.94, 1),
                line_width=1,
                padding=dp(16),
                spacing=dp(8),
            )
            empty_card.add_widget(MDBoxLayout())
            empty_card.add_widget(MDIcon(icon='archive-off-outline', theme_text_color='Hint', font_size='38sp', halign='center', pos_hint={'center_x': 0.5}))
            empty_card.add_widget(MDLabel(text='Aucune sauvegarde trouvée dans MagPro_Backups.', halign='center', font_style='Subtitle2', bold=True, theme_text_color='Secondary', adaptive_height=True))
            empty_card.add_widget(MDBoxLayout())
            lst.add_widget(empty_card)
        else:
            for fp, fn, mtime, sz in backups:
                is_zip = fn.lower().endswith('.zip')
                dt_str = datetime.fromtimestamp(mtime).strftime('%d/%m/%Y %H:%M')
                ico_col = (0.45, 0.2, 0.72, 1) if is_zip else (0.0, 0.52, 0.62, 1)
                ico_bg = (0.95, 0.92, 1, 1) if is_zip else (0.90, 0.97, 0.99, 1)

                card = MDCard(
                    orientation='horizontal',
                    size_hint=(1, None),
                    height=dp(66),
                    radius=[12],
                    elevation=0,
                    md_bg_color=(1, 1, 1, 1),
                    line_color=(0.88, 0.90, 0.94, 1),
                    line_width=1,
                    padding=[dp(10), dp(8)],
                    spacing=dp(10),
                    ripple_behavior=True,
                    on_release=lambda x, p=fp: self._confirm_restore(p),
                )
                card.add_widget(
                    self._ui_icon_pill(
                        'folder-zip-outline' if is_zip else 'database-outline',
                        ico_bg,
                        ico_col,
                        size_dp=42,
                        radius=11,
                        icon_size='22sp',
                    )
                )
                info = MDBoxLayout(orientation='vertical', size_hint_x=1, adaptive_height=True, spacing=dp(2), pos_hint={'center_y': 0.5})
                info.add_widget(
                    MDLabel(
                        text=fn,
                        font_style='Subtitle2',
                        bold=True,
                        theme_text_color='Primary',
                        shorten=True,
                        shorten_from='center',
                        adaptive_height=True,
                    )
                )
                info.add_widget(
                    MDLabel(
                        text=f'{dt_str}  •  {_fmt_size(sz)}',
                        font_style='Caption',
                        theme_text_color='Secondary',
                        adaptive_height=True,
                    )
                )
                card.add_widget(info)
                card.add_widget(
                    MDIcon(
                        icon='backup-restore',
                        theme_text_color='Custom',
                        text_color=(0.85, 0.45, 0.0, 1),
                        font_size='21sp',
                        size_hint=(None, None),
                        size=(dp(24), dp(24)),
                        pos_hint={'center_y': 0.5},
                    )
                )
                lst.add_widget(card)

        scroll.add_widget(lst)
        content.add_widget(scroll)

        btns = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint=(1, None), height=dp(48))
        btns.add_widget(
            self._ui_card_btn(
                'ANNULER',
                (0.92, 0.93, 0.95, 1),
                text_color=(0.35, 0.35, 0.4, 1),
                size_hint_x=0.36,
                font_size='12.5sp',
                on_release=lambda x: self.restore_dialog.dismiss() if getattr(self, 'restore_dialog', None) else None,
            )
        )
        btns.add_widget(
            self._ui_card_btn(
                'AUTRE FICHIER',
                (0.09, 0.44, 0.84, 1),
                icon='folder-open-outline',
                size_hint_x=0.64,
                font_size='12.5sp',
                elevation=1,
                on_release=lambda x: self._open_external_restore_picker(),
            )
        )
        content.add_widget(btns)

        self.restore_dialog = MDDialog(
            title='Restaurer (MagPro_Backups)',
            type='custom',
            content_cls=content,
            size_hint=(0.95, None),
            radius=[18, 18, 18, 18],
            buttons=[],
        )
        self._fit_dialog(self.restore_dialog, top_pad=16, side_pad=14, bottom_pad=12, spacer_top=8)
        self.restore_dialog.open()
        Clock.schedule_once(lambda dt: setattr(scroll, 'scroll_y', 1.0), 0.05)

    def _open_external_restore_picker(self):
        if getattr(self, 'restore_dialog', None):
            self.restore_dialog.dismiss()
            self.restore_dialog = None

        if platform == 'android':
            try:
                from android import activity
                from jnius import autoclass, cast

                intent_cls = autoclass('android.content.Intent')
                intent = intent_cls(intent_cls.ACTION_GET_CONTENT)
                intent.setType('*/*')
                intent.addCategory(intent_cls.CATEGORY_OPENABLE)
                activity.bind(on_activity_result=self._on_restore_file_chosen)
                cast('android.app.Activity', PythonActivity.mActivity).startActivityForResult(intent, 102)
            except Exception as exc:
                self.notify(f'Erreur sélecteur : {exc}', 'error')
            return

        default_path = self.get_backup_directory()
        content = MDBoxLayout(orientation='vertical', spacing=10, size_hint_y=None, height=dp(520))
        bg = MDCard(orientation='vertical', size_hint_y=1, md_bg_color=(0.15, 0.15, 0.15, 1), radius=[10], padding='5dp')
        self.file_chooser = FileChooserListView(path=default_path, filters=['*.zip', '*.db'], size_hint_y=1)
        bg.add_widget(self.file_chooser)
        content.add_widget(bg)
        btns = MDBoxLayout(spacing=10, size_hint_y=None, height=dp(50))
        btns.add_widget(self._ui_card_btn('ANNULER', (0.92, 0.93, 0.95, 1), text_color=(0.35, 0.35, 0.4, 1), size_hint_x=0.4, on_release=lambda x: self.restore_dialog.dismiss()))
        btns.add_widget(self._ui_card_btn('RESTAURER', (0.8, 0, 0, 1), icon='backup-restore', size_hint_x=0.6, elevation=1, on_release=lambda x: self._confirm_restore(self.file_chooser.selection[0] if self.file_chooser.selection else None)))
        content.add_widget(btns)
        self.restore_dialog = MDDialog(title='Restaurer une sauvegarde', type='custom', content_cls=content, size_hint=(0.95, None), radius=[18, 18, 18, 18])
        self._fit_dialog(self.restore_dialog)
        self.restore_dialog.open()

    def _confirm_restore(self, final_path: str | None):
        if getattr(self, '_restore_in_progress', False):
            return
        if not final_path or not os.path.exists(final_path):
            self.notify('Veuillez sélectionner un fichier valide.', 'error')
            return
        if getattr(self, 'restore_dialog', None):
            self.restore_dialog.dismiss()
            self.restore_dialog = None

        def _do_restore():
            if getattr(self, '_restore_in_progress', False):
                return
            self._restore_in_progress = True
            self._show_restore_progress_dialog(os.path.basename(final_path))
            threading.Thread(target=self._restore_worker, args=(final_path,), daemon=True).start()

        self.final_restore_confirm = self._ui_confirm_dialog(
            title='Attention',
            message=f"Voulez-vous vraiment restaurer '{os.path.basename(final_path)}' ?\nToutes les données actuelles seront remplacées.",
            confirm_text='OUI, RESTAURER',
            icon='backup-restore',
            confirm_icon='backup-restore',
            on_confirm=lambda: Clock.schedule_once(lambda dt: _do_restore(), 0.22),
        )

    def _show_restore_progress_dialog(self, filename: str):
        from kivymd.uix.progressbar import MDProgressBar

        if getattr(self, 'restore_progress_dialog', None):
            try:
                self.restore_progress_dialog.dismiss()
            except Exception:
                pass

        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(12),
            size_hint_y=None,
            adaptive_height=True,
            padding=[0, dp(4), 0, dp(4)],
        )
        badge = self._ui_icon_pill(
            'backup-restore',
            (1, 0.95, 0.88, 1),
            (0.85, 0.45, 0.0, 1),
            size_dp=56,
            radius=28,
            icon_size='28sp',
        )
        badge.pos_hint = {'center_x': 0.5}
        content.add_widget(badge)

        content.add_widget(
            MDLabel(
                text='Restauration en cours...',
                halign='center',
                font_style='H6',
                bold=True,
                theme_text_color='Primary',
                size_hint_y=None,
                height=dp(28),
            )
        )
        content.add_widget(
            MDLabel(
                text=filename,
                halign='center',
                font_style='Caption',
                bold=True,
                theme_text_color='Secondary',
                shorten=True,
                shorten_from='center',
                size_hint_y=None,
                height=dp(18),
            )
        )

        bar_card = MDCard(
            orientation='vertical',
            size_hint=(1, None),
            adaptive_height=True,
            radius=[12],
            elevation=0,
            md_bg_color=(0.96, 0.97, 0.99, 1),
            line_color=(0.86, 0.89, 0.94, 1),
            line_width=1,
            padding=[dp(12), dp(12)],
            spacing=dp(8),
        )
        progress_bar = MDProgressBar(
            max=100,
            value=0,
            size_hint_y=None,
            height=dp(8),
            color=(0.85, 0.45, 0.0, 1),
            back_color=(0.86, 0.89, 0.94, 1),
        )
        bar_card.add_widget(progress_bar)

        pct_row = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(20))
        lbl_pct = MDLabel(
            text='0%',
            halign='left',
            font_style='Subtitle2',
            bold=True,
            theme_text_color='Custom',
            text_color=(0.85, 0.45, 0.0, 1),
            size_hint_x=0.30,
        )
        lbl_step = MDLabel(
            text='Préparation...',
            halign='right',
            font_style='Caption',
            bold=True,
            theme_text_color='Secondary',
            shorten=True,
            size_hint_x=0.70,
        )
        pct_row.add_widget(lbl_pct)
        pct_row.add_widget(lbl_step)
        bar_card.add_widget(pct_row)
        content.add_widget(bar_card)

        self._restore_progress_widgets = {
            'bar': progress_bar,
            'lbl_pct': lbl_pct,
            'lbl_step': lbl_step,
        }

        self.restore_progress_dialog = MDDialog(
            title='',
            type='custom',
            content_cls=content,
            size_hint=(0.90, None),
            auto_dismiss=False,
            radius=[20, 20, 20, 20],
            buttons=[],
        )
        self._fit_dialog(self.restore_progress_dialog, top_pad=16, side_pad=14, bottom_pad=14, spacer_top=0)
        self.restore_progress_dialog.open()

    def _update_restore_progress_ui(self, pct: int, step_text: str):
        widgets = getattr(self, '_restore_progress_widgets', None)
        if not widgets or not getattr(self, 'restore_progress_dialog', None):
            return
        pct = max(0, min(100, int(pct)))
        widgets['bar'].value = pct
        widgets['lbl_pct'].text = f'{pct}%'
        widgets['lbl_step'].text = step_text

    def _restore_worker(self, final_path: str):
        temp_db_out = os.path.join(self.user_data_dir, f'temp_restore_stage_{int(time.time())}.db')
        last_ui_ts = 0.0

        def _push_ui(pct: int, text: str, force: bool = False):
            nonlocal last_ui_ts
            now_ts = time.monotonic()
            if force or (now_ts - last_ui_ts) >= 0.05:
                last_ui_ts = now_ts
                Clock.schedule_once(lambda dt, p=pct, t=text: self._update_restore_progress_ui(p, t), 0)

        try:
            _push_ui(5, 'Nettoyage WAL...', force=True)
            self.db.cleanup_wal_files()
            target_db = self.db.db_path
            base_dir_abs = os.path.abspath(self.user_data_dir)

            if final_path.lower().endswith('.zip'):
                with zipfile.ZipFile(final_path, 'r') as zf:
                    names = zf.namelist()
                    db_members = [n for n in names if n.lower().endswith('.db') and '..' not in n.replace('\\', '/').split('/')]
                    if not db_members:
                        raise ValueError('Archive invalide (aucune base .db trouvée).')

                    asset_members = [
                        n for n in names
                        if not n.endswith('/')
                        and (n.startswith('product_images/') or n.startswith('store_assets/'))
                        and not os.path.isabs(n)
                        and '..' not in n.replace('\\', '/').split('/')
                    ]
                    db_info = zf.getinfo(db_members[0])
                    total_bytes = max(1, db_info.file_size + sum(zf.getinfo(m).file_size for m in asset_members))
                    done_bytes = 0

                    _push_ui(8, 'Restauration base de données...', force=True)
                    with zf.open(db_members[0]) as src, open(temp_db_out, 'wb') as dst:
                        while True:
                            chunk = src.read(65536)
                            if not chunk:
                                break
                            dst.write(chunk)
                            done_bytes += len(chunk)
                            _push_ui(8 + int(82 * min(1.0, done_bytes / total_bytes)), 'Restauration base de données...')

                    os.replace(temp_db_out, target_db)

                    total_assets = len(asset_members)
                    for idx, member in enumerate(asset_members, start=1):
                        dest_path = os.path.abspath(os.path.join(self.user_data_dir, member))
                        if not dest_path.startswith(base_dir_abs):
                            continue
                        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                        with zf.open(member) as src, open(dest_path, 'wb') as dst:
                            while True:
                                chunk = src.read(65536)
                                if not chunk:
                                    break
                                dst.write(chunk)
                                done_bytes += len(chunk)
                        _push_ui(
                            8 + int(82 * min(1.0, done_bytes / total_bytes)),
                            f'Extraction images ({idx}/{total_assets})...',
                            force=(idx == total_assets),
                        )
            else:
                total_bytes = max(1, os.path.getsize(final_path))
                done_bytes = 0
                _push_ui(8, 'Copie base de données...', force=True)
                with open(final_path, 'rb') as src, open(temp_db_out, 'wb') as dst:
                    while True:
                        chunk = src.read(65536)
                        if not chunk:
                            break
                        dst.write(chunk)
                        done_bytes += len(chunk)
                        _push_ui(8 + int(82 * min(1.0, done_bytes / total_bytes)), 'Copie base de données...')
                os.replace(temp_db_out, target_db)

            _push_ui(95, 'Initialisation...', force=True)
            new_db = DatabaseManager()
            if os.path.basename(final_path).startswith('temp_restore') and os.path.abspath(os.path.dirname(final_path)) == base_dir_abs:
                try:
                    os.remove(final_path)
                except Exception:
                    pass
            _push_ui(100, 'Terminé !', force=True)
            Clock.schedule_once(lambda dt: self._finish_restore_ui(new_db, None), 0.12)
        except Exception as exc:
            if os.path.exists(temp_db_out):
                try:
                    os.remove(temp_db_out)
                except Exception:
                    pass
            Clock.schedule_once(lambda dt, e=str(exc): self._finish_restore_ui(None, e), 0)

    def _finish_restore_ui(self, new_db, error_msg: str | None):
        if getattr(self, 'restore_progress_dialog', None):
            try:
                self.restore_progress_dialog.dismiss()
            except Exception:
                pass
            self.restore_progress_dialog = None
        self._restore_progress_widgets = None

        if error_msg or new_db is None:
            self._restore_in_progress = False
            self.notify(f'Erreur restauration : {error_msg or "échec"}', 'error')
            return

        self.db = new_db
        self.is_seller_mode = self.db.get_bool_setting('config_seller_mode', default=False)
        self.cart = []
        self.selected_entity = None
        self.editing_tx_id = None
        self.editing_doc_type = None
        self.editing_payment_amount = Decimal('0.00')
        self.editing_payment_method = AppConstants.PAY_DEFAULT
        self.editing_global_discount = Decimal('0.00')
        self.current_editing_date = None

        def _apply_dash_after_close(dt):
            try:
                if getattr(self, 'sm', None):
                    self.sm.current = 'dashboard'
                self.update_dashboard_layout()
                self.notify('Restauration réussie !', 'success')
            finally:
                self._restore_in_progress = False

        Clock.schedule_once(_apply_dash_after_close, 0.24)

    def _on_restore_file_chosen(self, req_code, res_code, intent):
        from android import activity
        activity.unbind(on_activity_result=self._on_restore_file_chosen)
        if req_code == 102 and res_code == -1 and intent and intent.getData():
            uri = intent.getData()
            threading.Thread(target=self._copy_restore_uri_worker, args=(uri,), daemon=True).start()

    def _copy_restore_uri_worker(self, uri):
        try:
            from jnius import autoclass, cast
            cur_act = cast('android.app.Activity', PythonActivity.mActivity)
            resolver = cur_act.getContentResolver()
            ext = '.zip'
            try:
                oc = autoclass('android.provider.OpenableColumns')
                c = resolver.query(uri, None, None, None, None)
                if c and c.moveToFirst():
                    idx = c.getColumnIndex(oc.DISPLAY_NAME)
                    if idx >= 0 and str(c.getString(idx)).lower().endswith('.db'):
                        ext = '.db'
                    c.close()
            except Exception:
                pass
            temp_path = os.path.join(self.user_data_dir, f'temp_restore{ext}')
            stream = resolver.openInputStream(uri)
            with open(temp_path, 'wb') as out:
                buf = bytearray(8192)
                while True:
                    n = stream.read(buf)
                    if n == -1:
                        break
                    out.write(buf[:n])
            stream.close()
            Clock.schedule_once(lambda dt: self._confirm_restore(temp_path), 0)
        except Exception as exc:
            Clock.schedule_once(lambda dt, e=exc: self.notify(f'Erreur lecture fichier : {e}', 'error'), 0)

    # --------------------------------------------------------------------------
    # 7.14 NAVIGATION
    # --------------------------------------------------------------------------
    def go_back(self):
        try:
            Window.release_all_keyboards()
            self.editing_tx_id = None
            self.editing_doc_type = None
            self.editing_payment_amount = Decimal('0.00')
            self.editing_payment_method = AppConstants.PAY_DEFAULT
            self.editing_global_discount = Decimal('0.00')
            self.current_editing_date = None
            if hasattr(self, 'search_field') and self.search_field:
                self.search_field.set_raw_value('')
            self.cart = []
            self.selected_entity = None
            self.update_cart_button()
            self.sm.current = 'dashboard'
            self._reset_notification_state(0)
        except Exception:
            self.sm.current = 'dashboard'


# ==============================================================================
# 8. RAPPORT D'ERREURS GLOBAL & POINT D'ENTRÉE
# ==============================================================================
def send_error_to_telegram(error_traceback, error_type='Erreur Système'):
    def _send_worker():
        BOT_TOKEN = "GITHUB_SECRET_TELEGRAM_BOT_TOKEN"
        MASTER_ADMIN_ID = "GITHUB_SECRET_TELEGRAM_ADMIN_ID"

        if "GITHUB_SECRET" in BOT_TOKEN:
            return

        device_id = 'UNKNOWN'
        if platform == 'android':
            try:
                from jnius import autoclass
                pa = autoclass('org.kivy.android.PythonActivity')
                content_resolver = pa.mActivity.getContentResolver()
                secure = autoclass('android.provider.Settings$Secure')
                device_id = str(secure.getString(content_resolver, secure.ANDROID_ID))
            except Exception:
                device_id = 'ANDROID_ERROR'
        else:
            device_id = 'PC_WINDOWS'

        time_now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        safe_traceback = error_traceback
        if len(safe_traceback) > 3500:
            safe_traceback = '... [TRUNCATED] ...\n' + safe_traceback[-3500:]

        msg = (
            f'🚨 **تنبيه خطأ برمجي (CRASH REPORT)** 🚨\n\n'
            f'📱 **الجهاز:** `{device_id}`\n'
            f'⏰ **الوقت:** `{time_now}`\n'
            f'📌 **نوع الخطأ:** {error_type}\n\n'
            f'🛠 **التفاصيل:**\n```python\n{safe_traceback}\n```'
        )
        try:
            import requests
            requests.post(
                f'https://api.telegram.org/bot{BOT_TOKEN}/sendMessage',
                json={'chat_id': MASTER_ADMIN_ID, 'text': msg, 'parse_mode': 'Markdown'},
                timeout=10,
            )
        except Exception:
            pass

    threading.Thread(target=_send_worker, daemon=True).start()


def console_excepthook(exc_type, exc_value, exc_tb):
    traceback.print_exception(exc_type, exc_value, exc_tb)
    err_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_tb))
    send_error_to_telegram(err_msg, error_type='Console Exception Hook')
    try:
        with open(os.path.join(APP_DIR, 'crash_log.txt'), 'w', encoding='utf-8') as f:
            f.write(err_msg)
    except Exception:
        pass


sys.excepthook = console_excepthook

if __name__ == '__main__':
    try:
        StockApp().run()
    except Exception:
        err_msg = traceback.format_exc()
        send_error_to_telegram(err_msg, error_type='App Critical Crash')
        try:
            with open(os.path.join(APP_DIR, 'crash_log.txt'), 'w', encoding='utf-8') as f:
                f.write(err_msg)
        except Exception:
            pass
        if platform == 'android':
            try:
                from jnius import autoclass
                pa = autoclass('org.kivy.android.PythonActivity')
                files_dir = pa.mActivity.getExternalFilesDir(None).getAbsolutePath()
                with open(os.path.join(files_dir, 'magpro_crash.txt'), 'w', encoding='utf-8') as f:
                    f.write(err_msg)
            except Exception:
                pass



