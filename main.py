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
    MDFillRoundFlatButton,
    MDFillRoundFlatIconButton,
    MDFlatButton,
    MDIconButton,
    MDRaisedButton,
    MDRoundFlatIconButton,
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
    IconLeftWidget,
    IconRightWidget,
    MDList,
    OneLineAvatarIconListItem,
    OneLineListItem,
    TwoLineAvatarIconListItem,
)
from kivymd.uix.pickers import MDDatePicker
from kivymd.uix.screen import MDScreen
from kivymd.uix.screenmanager import MDScreenManager
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.selectioncontrol import MDCheckbox
from kivymd.uix.snackbar import Snackbar
from kivymd.uix.textfield import MDTextField
from kivymd.uix.toolbar import MDTopAppBar

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

# Moteur de mise en forme arabe (référence MagPro Vendeur)
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
    """Met en forme le texte arabe pour le rendu visuel (données utilisateur)."""
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

    # Limites strictes (parité MagPro Desktop)
    UNLIMITED_STOCK_VALUE = -999999999999.0
    UNLIMITED_STOCK_THRESHOLD = -900000000000.0
    MAX_AMOUNT_LIMIT = Decimal('1000000000')
    MAX_QTY_LIMIT = Decimal('100000000')
    MAX_STOCK_DB_LIMIT = Decimal('9999999999999')
    MAX_PERCENT_LIMIT = Decimal('100')

    # Entités & Produit par défaut (constantes uniques)
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

    # Codes documents uniques
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
        'Loyer',
        'Électricité / Gaz / Eau',
        'Salaires / Avances',
        'Transport / Carburant',
        'Maintenance',
        'Fournitures',
        'Autre',
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
        """
        DÉCISION UNIQUE DE CALCUL :
        Identique ligne pour ligne à `calculate_document_totals` du logiciel Desktop.
        """
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
                'allow_negative_stock': 'False',
                'warn_price_below_cost': 'True',
                'printer_auto': 'False',
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

    # --- Numérotation des documents (parité Desktop) ---
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
                        activity, rc, nif, nis, nai, price_category, gps_location
                    ) VALUES (?, ?, ?, 0, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''',
                    (e_type,) + fields,
                )
                e_id = cur.lastrowid
            else:
                cur.execute(
                    '''
                    UPDATE entities SET
                        name=?, phone=?, address=?, email=?, activity=?,
                        rc=?, nif=?, nis=?, nai=?, price_category=?, gps_location=?
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
                    running_global += qty
                    if ev['loc'] == AppConstants.LOC_WAREHOUSE:
                        running_wh += qty
                    else:
                        running_store += qty
                elif direction == -1:
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
        if final_pmp > Decimal('0'):
            cursor.execute(
                '''
                UPDATE transaction_items
                SET cost_at_time = ?
                WHERE product_id = ?
                  AND (cost_at_time IS NULL OR cost_at_time = 0)
                  AND transaction_id IN (
                      SELECT id FROM transactions WHERE doc_type IN (?, ?, ?)
                  )
                ''',
                (float(final_pmp), product_id, AppConstants.DOC_BV, AppConstants.DOC_FC, AppConstants.DOC_RC),
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

        if doc_type == AppConstants.DOC_BI and len({it.get('id') for it in items}) > 1:
            raise ValueError('Le Bon Initial (BI) est limité à un seul produit.')

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

            s_factor = AppConstants.STOCK_MOVEMENTS.get(doc_type, 0)
            allow_negative = self.get_bool_setting('allow_negative_stock', default=False)
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
        include_tx = doc_type_filter != AppConstants.DOC_TR
        include_tr = (
            entity_id is None
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
                if doc_type_filter != AppConstants.FAM_ALL:
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

    # --- Statistiques Financières (100% parité Desktop compute_financial_summary) ---
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
                        for it in items_rows:
                            q = FinancialEngine.to_decimal(it['qty'])
                            p = FinancialEngine.to_decimal(it['price'])
                            d_amt = FinancialEngine.to_decimal(it['discount_amount'])
                            u_disc = d_amt if it['discount_type'] == 'fixed' else p * (d_amt / Decimal('100'))
                            line_ht = FinancialEngine.quantize_money((p - u_disc) * q)
                            all_ht += line_ht
                            if not AppConstants.is_default_product(it['product_id']):
                                real_ht += line_ht
                                cogs += FinancialEngine.quantize_money(q * FinancialEngine.to_decimal(it['cost_at_time']))
                        if real_ht > Decimal('0'):
                            gd_ht = calc['total_ht_net_of_lines'] - calc['total_ht_final']
                            if all_ht > Decimal('0') and real_ht < all_ht:
                                gd_ht = FinancialEngine.quantize_money(gd_ht * (real_ht / all_ht))
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

    def header(self):
        if self.page_no() != 1:
            return
        store_name = self.smart_text(self.store_info.get('store_name', AppConstants.APP_NAME))
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

        doc_type = self.doc_info.get('doc_type')
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

        self.set_y(8)
        block_w = 120
        self.set_x((self.w - block_w) / 2)
        self.set_font('arabic', 'B', 22)
        self.cell(block_w, 10, store_name, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.set_x((self.w - block_w) / 2)
        self.set_font('arabic', '', 9)
        self.multi_cell(block_w, 5, '\n'.join(info_lines), align='C')
        self.ln(8)

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
    height: dp(90)
    padding: 0
    spacing: 0

    MDCard:
        orientation: 'horizontal'
        padding: dp(10)
        spacing: dp(10)
        radius: [8]
        elevation: 1
        ripple_behavior: True
        on_release: root.on_tap()
        md_bg_color: (1, 1, 1, 1)

        MDCard:
            size_hint: None, None
            size: dp(55), dp(55)
            radius: [8]
            elevation: 0
            padding: 0
            pos_hint: {'center_y': .5}
            md_bg_color: (0.95, 0.95, 0.95, 1)
            ripple_behavior: True
            on_release: root.on_zoom()

            MDFloatLayout:
                size_hint: 1, 1
                MDIcon:
                    icon: root.icon_name
                    theme_text_color: "Custom"
                    text_color: root.icon_color
                    pos_hint: {'center_x': .5, 'center_y': .5}
                    font_size: '30sp'
                    opacity: 0 if root.image_path else 1

                FitImage:
                    source: root.image_path if root.image_path else ''
                    radius: [8]
                    mipmap: True
                    opacity: 1 if root.image_path else 0
                    size_hint: 1, 1
                    pos_hint: {'center_x': .5, 'center_y': .5}
                    fit_mode: "cover"

        MDBoxLayout:
            orientation: 'vertical'
            pos_hint: {'center_y': .5}
            spacing: dp(5)

            MDLabel:
                text: root.text_name
                font_style: "Subtitle1"
                bold: True
                text_size: self.width, None
                max_lines: 2
                halign: 'left'
                valign: 'top'
                font_size: '16sp'
                theme_text_color: "Custom"
                text_color: (0.1, 0.1, 0.1, 1)
                font_name: 'ArabicFont'

            MDBoxLayout:
                orientation: 'horizontal'
                spacing: dp(5)

                MDLabel:
                    text: root.text_price
                    font_style: "H6"
                    theme_text_color: "Custom"
                    text_color: root.price_color
                    bold: True
                    size_hint_x: None
                    adaptive_width: True
                    font_size: '19sp'
                    font_name: 'ArabicFont'
                    max_lines: 1

                MDLabel:
                    text: root.text_stock
                    theme_text_color: "Custom"
                    text_color: (0.1, 0.1, 0.1, 1)
                    halign: 'right'
                    size_hint_x: 1
                    bold: True
                    font_size: '13sp'
                    max_lines: 1
                    shorten: True
                    font_name: 'ArabicFont'

<ProductRecycleView>:
    viewclass: 'ProductRecycleItem'
    RecycleBoxLayout:
        default_size: None, dp(95)
        default_size_hint: 1, None
        size_hint_y: None
        height: self.minimum_height
        orientation: 'vertical'
        spacing: dp(4)
        padding: dp(5)

<HistoryRecycleItem>:
    orientation: "vertical"
    size_hint_y: None
    height: dp(130)
    padding: dp(12)
    spacing: dp(6)
    radius: [14]
    elevation: 2
    ripple_behavior: True
    md_bg_color: root.bg_color
    on_release: root.on_tap()

    MDBoxLayout:
        orientation: "horizontal"
        size_hint_y: None
        height: dp(38)
        spacing: dp(10)

        MDCard:
            size_hint: None, None
            size: dp(38), dp(38)
            radius: [19]
            md_bg_color: (0.95, 0.95, 0.95, 1)
            elevation: 0
            pos_hint: {"center_y": .5}

            MDIcon:
                icon: root.icon_name
                theme_text_color: "Custom"
                text_color: root.icon_color
                pos_hint: {"center_x": .5, "center_y": .5}
                font_size: "22sp"

        MDBoxLayout:
            orientation: "vertical"
            spacing: dp(2)
            pos_hint: {"center_y": .5}

            MDLabel:
                text: root.text_type
                bold: True
                font_style: "Subtitle1"
                theme_text_color: "Primary"
                font_name: 'ArabicFont'
                halign: "left"

            MDLabel:
                text: root.text_ref
                font_style: "Caption"
                theme_text_color: "Secondary"
                font_name: 'ArabicFont'
                halign: "left"

    MDLabel:
        text: root.text_entity
        font_style: "Subtitle1"
        font_size: "16sp"
        bold: True
        theme_text_color: "Custom"
        text_color: (0.1, 0.1, 0.1, 1)
        font_name: 'ArabicFont'
        size_hint_y: None
        height: self.texture_size[1]
        text_size: self.width, None
        max_lines: 2
        halign: "left"

    MDBoxLayout:
        size_hint_y: 1

    MDBoxLayout:
        orientation: "horizontal"
        size_hint_y: None
        height: dp(24)

        MDLabel:
            text: root.text_date
            font_style: "Caption"
            theme_text_color: "Hint"
            halign: "left"
            size_hint_x: 0.4
            font_name: 'ArabicFont'

        MDLabel:
            text: root.text_amount
            font_style: "Subtitle1"
            bold: True
            theme_text_color: "Custom"
            text_color: root.icon_color
            halign: "right"
            font_size: "17sp"
            size_hint_x: 0.6
            font_name: 'ArabicFont'

<HistoryRecycleView>:
    viewclass: 'HistoryRecycleItem'
    RecycleBoxLayout:
        default_size: None, dp(136)
        default_size_hint: 1, None
        size_hint_y: None
        height: self.minimum_height
        orientation: 'vertical'
        spacing: dp(6)
        padding: dp(5)

<EntityRecycleItem>:
    orientation: "horizontal"
    size_hint_y: None
    height: dp(80)
    padding: dp(10)
    spacing: dp(15)
    ripple_behavior: True
    md_bg_color: (1, 1, 1, 1)
    radius: [0]
    on_release: root.on_tap()

    MDIcon:
        icon: root.icon_name
        theme_text_color: "Custom"
        text_color: root.icon_color
        pos_hint: {"center_y": .5}
        font_size: "32sp"
        size_hint_x: None
        width: dp(40)

    MDBoxLayout:
        orientation: "vertical"
        pos_hint: {"center_y": .5}
        size_hint_x: 1
        spacing: dp(4)

        MDLabel:
            text: root.text_name
            bold: True
            font_style: "Subtitle1"
            font_name: 'ArabicFont'
            theme_text_color: "Custom"
            text_color: (0.1, 0.1, 0.1, 1)
            shorten: True
            shorten_from: 'right'

        MDLabel:
            text: root.text_balance
            font_style: "Caption"
            font_name: 'ArabicFont'
            markup: True
            theme_text_color: "Secondary"

<EntityRecycleView>:
    viewclass: 'EntityRecycleItem'
    RecycleBoxLayout:
        default_size: None, dp(80)
        default_size_hint: 1, None
        size_hint_y: None
        height: self.minimum_height
        orientation: 'vertical'
        spacing: dp(2)
        padding: 0

<MgmtEntityRecycleItem>:
    orientation: "horizontal"
    size_hint_y: None
    height: dp(80)
    padding: dp(10)
    spacing: dp(5)
    ripple_behavior: True
    md_bg_color: root.bg_color
    on_release: root.on_pay()

    MDIcon:
        icon: "account-circle"
        theme_text_color: "Custom"
        text_color: (0.5, 0.5, 0.5, 1)
        pos_hint: {"center_y": .5}
        font_size: "32sp"
        size_hint_x: None
        width: dp(40)

    MDBoxLayout:
        orientation: "vertical"
        pos_hint: {"center_y": .5}
        size_hint_x: 1
        spacing: dp(2)
        padding: [dp(10), 0, 0, 0]

        MDLabel:
            text: root.text_name
            bold: True
            font_style: "Subtitle1"
            font_name: 'ArabicFont'
            theme_text_color: "Custom"
            text_color: (0.1, 0.1, 0.1, 1)
            shorten: True
            shorten_from: 'right'
            halign: "left"

        MDLabel:
            text: root.text_balance
            font_style: "Caption"
            font_name: 'ArabicFont'
            markup: True
            theme_text_color: "Secondary"
            halign: "left"

    MDIconButton:
        icon: "clock-time-eight-outline"
        theme_text_color: "Custom"
        text_color: (0, 0.5, 0.5, 1)
        pos_hint: {"center_y": .5}
        on_release: root.on_history()

<MgmtEntityRecycleView>:
    viewclass: 'MgmtEntityRecycleItem'
    RecycleBoxLayout:
        default_size: None, dp(80)
        default_size_hint: 1, None
        size_hint_y: None
        height: self.minimum_height
        orientation: 'vertical'
        spacing: dp(2)
        padding: 0

<CartRecycleItem>:
    orientation: "horizontal"
    size_hint_y: None
    height: dp(85)
    padding: [dp(15), 0, 0, 0]
    md_bg_color: 1, 1, 1, 1
    radius: [0]
    ripple_behavior: True
    on_release: root.on_tap()

    MDBoxLayout:
        orientation: "vertical"
        pos_hint: {"center_y": .5}
        adaptive_height: True
        spacing: dp(4)

        MDLabel:
            text: root.text_name
            font_style: "Subtitle1"
            bold: True
            theme_text_color: "Primary"
            adaptive_height: True
            font_name: 'ArabicFont'

        MDLabel:
            text: root.text_details
            font_size: "16sp"
            theme_text_color: "Custom"
            text_color: root.details_color
            bold: True
            adaptive_height: True
            font_name: 'ArabicFont'

    MDIconButton:
        icon: "delete"
        theme_text_color: "Custom"
        text_color: (0.9, 0, 0, 1)
        pos_hint: {"center_y": .5}
        icon_size: "24sp"
        on_release: root.on_delete()

<CartRecycleView>:
    viewclass: 'CartRecycleItem'
    RecycleBoxLayout:
        default_size: None, dp(85)
        default_size_hint: 1, None
        size_hint_y: None
        height: self.minimum_height
        orientation: 'vertical'
        spacing: dp(1)
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
    details_color = ListProperty([0.4, 0.4, 0.4, 1])
    raw_data = ObjectProperty(None, allownone=True)

    def refresh_view_attrs(self, rv, index, data):
        self.index = index
        self.text_name = ar_visual(data.get('name', ''))
        self.text_details = data.get('details', '')
        self.details_color = data.get('d_color', [0.4, 0.4, 0.4, 1])
        self.raw_data = data.get('raw_item')
        return super().refresh_view_attrs(rv, index, data)

    def on_tap(self):
        app = MDApp.get_running_app()
        if app and self.raw_data:
            app.edit_cart_item(self.raw_data)

    def on_delete(self):
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

    def on_touch_down(self, touch):
        if self.collide_point(*touch.pos):
            self._is_long_press = False
            self._start_pos = touch.pos
            self._long_press_event = Clock.schedule_once(lambda dt: self._trigger_long_press(), 0.45)
        return super().on_touch_down(touch)

    def on_touch_move(self, touch):
        if self._long_press_event:
            if abs(touch.x - self._start_pos[0]) > dp(20) or abs(touch.y - self._start_pos[1]) > dp(20):
                self._long_press_event.cancel()
                self._long_press_event = None
        return super().on_touch_move(touch)

    def on_touch_up(self, touch):
        if self._long_press_event:
            self._long_press_event.cancel()
            self._long_press_event = None
        if self._is_long_press:
            return True
        return super().on_touch_up(touch)

    def _trigger_long_press(self):
        self._is_long_press = True
        app = MDApp.get_running_app()
        if app and self.entity_data:
            app.open_entity_edit_menu(self.entity_data)

    def on_pay(self):
        if self._is_long_press:
            return
        app = MDApp.get_running_app()
        if app and self.entity_data:
            app.start_direct_payment_from_manager(self.entity_data)

    def on_history(self):
        app = MDApp.get_running_app()
        if app and self.entity_data:
            app.open_entity_history_dialog(self.entity_data)


class EntityRecycleItem(RecycleDataViewBehavior, MDCard):
    index = None
    text_name = StringProperty('')
    text_balance = StringProperty('')
    icon_name = StringProperty('account')
    icon_color = ListProperty([0, 0, 0, 1])
    entity_data = ObjectProperty(None, allownone=True)

    def refresh_view_attrs(self, rv, index, data):
        self.index = index
        self.text_name = ar_visual(data.get('raw_name', ''))
        self.text_balance = data.get('balance_text', '')
        self.icon_name = data.get('icon', 'account')
        self.icon_color = data.get('icon_color', [0, 0, 0, 1])
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
        self.do_scroll_x = False
        self.do_scroll_y = True
        self.scroll_type = ['bars', 'content']
        self.bar_width = dp(4)
        Clock.schedule_once(
            lambda dt: setattr(self.effect_y, 'friction', 0.12) if hasattr(self, 'effect_y') else None,
            0,
        )


class CartRecycleView(_SmoothRecycleView):
    pass


class ProductRecycleView(_SmoothRecycleView):
    def on_scroll_y(self, instance, value):
        app = MDApp.get_running_app()
        if app and value <= 0.15 and not self.loading_lock and not app.is_loading_more and len(self.data) >= 50:
            self.loading_lock = True
            app.load_more_products(reset=False)


class HistoryRecycleView(_SmoothRecycleView):
    def on_scroll_y(self, instance, value):
        app = MDApp.get_running_app()
        if app and value <= 0.1 and not self.loading_lock and not app.is_loading_history and len(self.data) >= 50:
            self.loading_lock = True
            app.load_more_history()


class EntityRecycleView(_SmoothRecycleView):
    def on_scroll_y(self, instance, value):
        app = MDApp.get_running_app()
        if app and value <= 0.1 and not self.loading_lock and not app.is_loading_entities and len(self.data) >= 50:
            self.loading_lock = True
            app.load_more_entities(reset=False)


class MgmtEntityRecycleView(_SmoothRecycleView):
    def on_scroll_y(self, instance, value):
        app = MDApp.get_running_app()
        if app and value <= 0.1 and not self.loading_lock and not app.is_loading_entities and len(self.data) >= 50:
            self.loading_lock = True
            app.load_more_entities(reset=False)


# ==============================================================================
# 7. APPLICATION PRINCIPALE (StockApp - MAGPRO MOBILE v8.1.0)
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

        self.history_page_offset = 0
        self.history_batch_size = 50
        self.is_loading_history = False
        self.history_start_date = datetime.now().date()
        self.history_end_date = datetime.now().date()
        self.current_history_filter = AppConstants.FAM_ALL

        self.entity_page_offset = 0
        self.is_loading_entities = False
        self.current_entity_type_mgmt = AppConstants.ENTITY_CLIENT
        self.current_entity_sort = 'name'
        self.active_entity_rv = None

        self._search_event = None
        self._entity_search_event = None
        self._notify_event = None
        self._last_click_time = 0
        self.is_showing_alert = False
        self.tone_gen = None
        self.db = None

        self.dialog = None
        self.rv_products = None
        self.status_bar_label = None
        self.status_bar_bg = None

    def fix_text(self, text) -> str:
        return ar_visual(text)

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

        self.status_bar_bg = MDCard(size_hint_y=None, height=dp(36), radius=[0], md_bg_color=(0.18, 0.18, 0.2, 1), elevation=0)
        self.status_bar_label = MDLabel(text='Prêt', halign='center', theme_text_color='Custom', text_color=(1, 1, 1, 1), font_style='Caption', bold=True)
        self.status_bar_bg.add_widget(self.status_bar_label)
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
        if not self.status_bar_label:
            return
        colors = {
            'success': (0, 0.6, 0.2, 1),
            'error': (0.8, 0.1, 0.1, 1),
            'warning': (0.9, 0.5, 0, 1),
            'info': (0.18, 0.18, 0.2, 1),
        }
        self.status_bar_label.text = str(text)
        self.status_bar_bg.md_bg_color = colors.get(msg_type, (0.18, 0.18, 0.2, 1))
        if self._notify_event:
            self._notify_event.cancel()
        self._notify_event = Clock.schedule_once(self._reset_notification_state, 3)

    def _reset_notification_state(self, dt):
        if self.status_bar_label:
            self._notify_event = None
            self.status_bar_label.text = 'Prêt'
            self.status_bar_bg.md_bg_color = (0.18, 0.18, 0.2, 1)

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
            for d_name in (
                'dialog', 'bt_dialog', 'zoom_dialog', 'img_picker_dialog', 'family_dialog',
                'add_fam_dialog', 'confirm_del_fam_dialog', 'settings_menu_dialog', 'filter_dialog',
                'mgmt_dialog', 'options_dialog', 'ae_dialog', 'cat_dialog', 'del_conf_dialog',
                'pass_dialog', 'logout_diag', 'store_settings_dialog', 'entity_dialog',
                'simple_pay_dialog', 'debt_dialog', 'overpay_dialog', 'pay_dialog',
                'srv_dialog', 'pending_dialog', 'entity_hist_dialog', 'auth_dialog',
                'toggle_dialog', 'mapping_diag', 'import_diag', 'restore_dialog',
                'final_restore_confirm', 'pdf_ready_dialog', 'expense_dialog',
                'history_exp_dialog', 'del_exp_dialog', 'edit_dialog', 'hist_filter_dialog',
                'stock_error_dialog', 'warn_cost_dlg', 'confirm_del_dialog',
            ):
                dlg = getattr(self, d_name, None)
                if dlg and dlg.parent:
                    dlg.dismiss()
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
        path = self._get_hidden_sys_file()
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return {'trial_start': None, 'attempts': 0}

    def save_persistent_data(self, data: dict):
        path = self._get_hidden_sys_file()
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f)
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
        content = MDBoxLayout(orientation='vertical', spacing='12dp', size_hint_y=None, adaptive_height=True, padding=['20dp', '20dp', '20dp', '10dp'])
        icon_name = 'shield-lock' if trial_expired else 'shield-check'
        icon_col = (0.8, 0, 0, 1) if trial_expired else self.theme_cls.primary_color
        content.add_widget(MDIcon(icon=icon_name, halign='center', font_size='64sp', theme_text_color='Custom', text_color=icon_col, pos_hint={'center_x': 0.5}))
        if trial_expired:
            t_txt = "Période d'essai terminée"
            s_txt = "L'application est verrouillée.\nVeuillez saisir votre clé d'activation pour continuer."
            t_col = 'Error'
        else:
            t_txt = f"Période d'essai : {days_left} jour(s) restant(s)"
            s_txt = "Profitez de votre essai gratuit ou activez l'application dès maintenant."
            t_col = 'Primary'
        content.add_widget(MDLabel(text=t_txt, halign='center', font_style='H5', bold=True, theme_text_color=t_col, adaptive_height=True))
        content.add_widget(MDLabel(text=s_txt, halign='center', font_style='Caption', theme_text_color='Secondary', adaptive_height=True))

        id_card = MDCard(orientation='vertical', radius=[10], padding=['15dp', '12dp'], md_bg_color=(0.96, 0.96, 0.96, 1), elevation=0, size_hint_y=None, adaptive_height=True, spacing='5dp')
        id_card.add_widget(MDLabel(text="ID de l'appareil :", halign='left', font_style='Caption', theme_text_color='Secondary', adaptive_height=True))
        id_row = MDBoxLayout(orientation='horizontal', spacing='10dp', adaptive_height=True)
        id_row.add_widget(MDTextField(text=device_id, readonly=True, font_size='16sp', mode='line', active_line=False, size_hint_x=0.85, pos_hint={'center_y': 0.5}))
        id_row.add_widget(MDIconButton(icon='content-copy', theme_text_color='Custom', text_color=self.theme_cls.primary_color, on_release=lambda x: [Clipboard.copy(device_id), self.notify('ID copié', 'success')], pos_hint={'center_y': 0.5}))
        id_card.add_widget(id_row)
        content.add_widget(id_card)

        key_row = MDBoxLayout(orientation='horizontal', spacing='10dp', adaptive_height=True)
        self.field_key = NoMenuTextField(hint_text='Saisir la clé de licence', mode='rectangle', size_hint_x=0.85, pos_hint={'center_y': 0.5})
        key_row.add_widget(self.field_key)
        key_row.add_widget(MDIconButton(icon='content-paste', theme_text_color='Custom', text_color=self.theme_cls.primary_color, on_release=lambda x: setattr(self.field_key, 'text', Clipboard.paste()), pos_hint={'center_y': 0.5}))
        content.add_widget(key_row)

        content.add_widget(
            MDRaisedButton(
                text="ACTIVER L'APPLICATION",
                md_bg_color=(0, 0.7, 0, 1),
                font_size='16sp',
                size_hint_x=1,
                size_hint_y=None,
                height='52dp',
                on_release=lambda x: self.validate_activation(self.field_key.text, self.activation_dialog_ref),
            )
        )
        self.online_act_btn = MDFillRoundFlatIconButton(
            icon='earth',
            text="Demander l'activation en ligne",
            md_bg_color=(0.1, 0.5, 0.8, 1),
            theme_text_color='Custom',
            text_color=(1, 1, 1, 1),
            icon_color=(1, 1, 1, 1),
            font_size='15sp',
            size_hint_x=1,
            size_hint_y=None,
            height='45dp',
            on_release=self.request_online_activation_mobile,
        )
        content.add_widget(self.online_act_btn)
        if not trial_expired:
            content.add_widget(
                MDFlatButton(
                    text="Continuer l'essai gratuit",
                    theme_text_color='Custom',
                    text_color=self.theme_cls.primary_color,
                    size_hint_x=1,
                    size_hint_y=None,
                    height='40dp',
                    on_release=lambda x: self.continue_trial(self.activation_dialog_ref),
                )
            )
        self.activation_dialog_ref = MDDialog(title='', type='custom', content_cls=content, size_hint=(0.9, None), auto_dismiss=False, radius=[16, 16, 16, 16])
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
        p_data['attempts'] = attempts + 1
        self.save_persistent_data(p_data)
        self.online_act_btn.text = 'Vérification...'
        self.online_act_btn.icon = 'sync'
        self.online_act_btn.disabled = True
        threading.Thread(target=self._online_activation_worker_mobile, daemon=True).start()

    def _online_activation_worker_mobile(self):
        BOT_TOKEN = "GITHUB_SECRET_TELEGRAM_BOT_TOKEN"
        MASTER_ADMIN_ID = "GITHUB_SECRET_TELEGRAM_ADMIN_ID"
        device_id = self.get_device_id()
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
            requests.post(
                f'https://api.telegram.org/bot{BOT_TOKEN}/sendMessage',
                json={'chat_id': MASTER_ADMIN_ID, 'text': msg, 'parse_mode': 'Markdown', 'reply_markup': reply_markup},
                timeout=10,
            )
        except Exception:
            def no_server(dt):
                self.notify("❌ Échec d'envoi au serveur", 'error')
                self.online_act_btn.text = "Demander l'activation en ligne"
                self.online_act_btn.icon = 'earth'
                self.online_act_btn.disabled = False
            Clock.schedule_once(no_server, 0)
            return

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
            height=dp(360),
            pos_hint={'center_x': 0.5, 'center_y': 0.52},
            padding=[dp(25), dp(30), dp(25), dp(25)],
            spacing=dp(15),
            radius=[24, 24, 24, 24],
            elevation=3,
            md_bg_color=(1, 1, 1, 1),
        )
        icon_box = MDFloatLayout(size_hint_y=None, height=dp(65))
        icon_box.add_widget(MDIcon(icon='store-check', font_size='58sp', pos_hint={'center_x': 0.5, 'center_y': 0.5}, theme_text_color='Custom', text_color=(0.1, 0.45, 0.85, 1)))
        card.add_widget(icon_box)
        card.add_widget(MDLabel(text=AppConstants.APP_NAME, halign='center', font_style='H5', bold=True, adaptive_height=True))
        card.add_widget(MDLabel(text='Authentification', halign='center', font_style='Caption', theme_text_color='Secondary', adaptive_height=True))

        self.password_field = SmartTextField(hint_text='Mot de passe', password=True, icon_right='lock-outline', halign='center')
        card.add_widget(self.password_field)
        card.add_widget(
            MDFillRoundFlatButton(
                text='SE CONNECTER',
                font_size='16sp',
                size_hint_x=1,
                size_hint_y=None,
                height=dp(50),
                md_bg_color=(0.1, 0.45, 0.85, 1),
                on_release=self.do_login,
            )
        )
        layout.add_widget(card)

        layout.add_widget(
            MDFlatButton(
                text='Changer / Définir le mot de passe',
                pos_hint={'center_x': 0.5, 'y': 0.1},
                theme_text_color='Custom',
                text_color=(0.35, 0.35, 0.4, 1),
                on_release=self.show_change_password_dialog_login,
            )
        )
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
        content = MDBoxLayout(orientation='vertical', spacing='15dp', size_hint_y=None, height=dp(260), padding=[0, dp(10), 0, 0])
        old_pass = SmartTextField(hint_text='Mot de passe actuel', password=True)
        new_pass = SmartTextField(hint_text='Nouveau mot de passe', password=True)
        conf_pass = SmartTextField(hint_text='Confirmer le mot de passe', password=True)
        content.add_widget(old_pass)
        content.add_widget(new_pass)
        content.add_widget(conf_pass)

        def _confirm(x):
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

        self.pass_dialog = MDDialog(
            title='Changer le mot de passe',
            type='custom',
            content_cls=content,
            buttons=[
                MDFlatButton(text='ANNULER', on_release=lambda x: self.pass_dialog.dismiss()),
                MDRaisedButton(text='CONFIRMER', md_bg_color=(0, 0.6, 0.2, 1), on_release=_confirm),
            ],
        )
        self.pass_dialog.open()

    def logout(self):
        def _do_exit(x):
            if getattr(self, 'logout_diag', None):
                self.logout_diag.dismiss()
            self.stop()

        self.logout_diag = MDDialog(
            title='Fermeture',
            text="Voulez-vous vraiment fermer l'application ?",
            buttons=[
                MDFlatButton(text='ANNULER', on_release=lambda x: self.logout_diag.dismiss()),
                MDRaisedButton(text='FERMER', md_bg_color=(0.8, 0, 0, 1), text_color=(1, 1, 1, 1), on_release=_do_exit),
            ],
        )
        self.logout_diag.open()

    # --------------------------------------------------------------------------
    # 7.3 TABLEAU DE BORD & STATISTIQUES FINANCIÈRES
    # --------------------------------------------------------------------------
    def _build_dashboard_screen(self):
        screen = MDScreen(name='dashboard')
        layout = MDBoxLayout(orientation='vertical')
        self.dash_toolbar = MDTopAppBar(
            title='Accueil',
            anchor_title='center',
            left_action_items=[['clipboard-text-clock', lambda x: self.show_pending_dialog()]],
            right_action_items=[['cog', lambda x: self.open_settings_menu()], ['logout', lambda x: self.logout()]],
        )
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
            elevation=2,
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
            padding=dp(15),
            radius=[14],
            ripple_behavior=True,
            on_release=action,
            md_bg_color=bg_color,
            elevation=2,
            size_hint_y=None,
            height=dp(108),
        )
        card.add_widget(MDIcon(icon=icon, font_size='36sp', pos_hint={'center_x': 0.5}, theme_text_color='Custom', text_color=icon_color))
        card.add_widget(MDLabel(text=text, halign='center', bold=True, font_style='Caption'))
        return card

    def update_dashboard_layout(self):
        if not getattr(self, 'buttons_container', None) or not getattr(self, 'stats_card_container', None):
            return
        self.buttons_container.clear_widgets()

        c_bv = ((0.88, 0.97, 0.9, 1), (0.1, 0.6, 0.2, 1))
        c_ba = ((1.0, 0.94, 0.84, 1), (0.9, 0.4, 0.0, 1))
        c_fc = ((0.88, 0.94, 1.0, 1), (0.0, 0.3, 0.7, 1))
        c_ff = ((0.98, 0.88, 0.88, 1), (0.7, 0.1, 0.2, 1))
        c_fp = ((0.94, 0.88, 1.0, 1), (0.5, 0.1, 0.6, 1))
        c_dp = ((0.85, 0.97, 0.97, 1), (0.0, 0.5, 0.5, 1))
        c_bs = ((0.95, 0.92, 0.88, 1), (0.5, 0.3, 0.1, 1))
        c_be = ((0.9, 0.9, 0.98, 1), (0.2, 0.2, 0.7, 1))
        c_rc = ((1.0, 0.85, 0.85, 1), (0.9, 0.0, 0.0, 1))
        c_rf = ((1.0, 0.88, 0.95, 1), (0.8, 0.1, 0.5, 1))
        c_tr = ((0.92, 0.94, 0.96, 1), (0.3, 0.4, 0.5, 1))
        c_pr = ((1.0, 0.97, 0.85, 1), (0.7, 0.6, 0.0, 1))
        c_cl = ((0.85, 0.95, 1.0, 1), (0.0, 0.6, 0.8, 1))
        c_fo = ((0.96, 0.93, 0.9, 1), (0.6, 0.4, 0.2, 1))
        c_de = ((0.99, 0.88, 0.88, 1), (0.8, 0.0, 0.0, 1))

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
        self.stats_card_container.spacing = dp(12)

        header_box = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True, spacing=dp(10))
        self.lbl_dashboard_date = MDLabel(text="Tableau de Bord (Aujourd'hui)", font_style='H6', bold=True, theme_text_color='Primary', halign='center', size_hint_y=None, height=dp(28))
        header_box.add_widget(self.lbl_dashboard_date)

        tabs_wrapper = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(36), spacing=dp(5), adaptive_width=True, pos_hint={'center_x': 0.5})
        self.btn_stat_today = MDRaisedButton(text='AUJ', size_hint_x=None, width=dp(54), font_size='11sp', elevation=0, on_release=lambda x: self.filter_dashboard_stats('auj'))
        self.btn_stat_yesterday = MDRaisedButton(text='HIE', size_hint_x=None, width=dp(54), font_size='11sp', elevation=0, md_bg_color=(0.5, 0.5, 0.5, 1), on_release=lambda x: self.filter_dashboard_stats('hier'))
        self.btn_stat_week = MDRaisedButton(text='SEM', size_hint_x=None, width=dp(54), font_size='11sp', elevation=0, md_bg_color=(0.5, 0.5, 0.5, 1), on_release=lambda x: self.filter_dashboard_stats('semaine'))
        self.btn_stat_month = MDRaisedButton(text='MOI', size_hint_x=None, width=dp(54), font_size='11sp', elevation=0, md_bg_color=(0.5, 0.5, 0.5, 1), on_release=lambda x: self.filter_dashboard_stats('mois'))
        self.btn_stat_date = MDRaisedButton(text='CAL', size_hint_x=None, width=dp(54), font_size='11sp', elevation=0, md_bg_color=(0.5, 0.5, 0.5, 1), on_release=self.open_stats_date_picker)
        for b in (self.btn_stat_today, self.btn_stat_yesterday, self.btn_stat_week, self.btn_stat_month, self.btn_stat_date):
            tabs_wrapper.add_widget(b)
        header_box.add_widget(tabs_wrapper)
        self.stats_card_container.add_widget(header_box)

        def _stat_card(icon: str, title: str, attr_name: str, bg_c, ic_c):
            c = MDCard(orientation='vertical', padding=dp(12), radius=[14], md_bg_color=bg_c, elevation=0, size_hint_y=None, height=dp(100))
            top = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(26), spacing=dp(6))
            top.add_widget(MDIcon(icon=icon, theme_text_color='Custom', text_color=ic_c, font_size='22sp', size_hint_x=None, width=dp(26)))
            top.add_widget(MDLabel(text=title, halign='left', font_style='Caption', bold=True, theme_text_color='Secondary'))
            lbl = MDLabel(text='0,00 DA', halign='center', font_style='H6', bold=True, theme_text_color='Custom', text_color=(0.15, 0.15, 0.15, 1))
            setattr(self, attr_name, lbl)
            c.add_widget(top)
            c.add_widget(MDBoxLayout(size_hint_y=None, height=dp(8)))
            c.add_widget(lbl)
            return c

        if self.is_seller_mode:
            self.stats_more_container = None
            stats_grid = MDGridLayout(cols=2, spacing=dp(10), adaptive_height=True)
            stats_grid.add_widget(_stat_card('cart-outline', 'Ventes Nettes', 'lbl_stat_sales', (0.9, 0.95, 1, 1), (0, 0.35, 0.8, 1)))
            stats_grid.add_widget(_stat_card('cash-plus', 'Encaissements', 'lbl_stat_in', (0.92, 0.98, 0.92, 1), (0, 0.6, 0.2, 1)))
            self.stats_card_container.add_widget(stats_grid)
        else:
            caisse_card = MDCard(
                radius=[16],
                md_bg_color=(0.98, 0.99, 0.98, 1),
                line_color=(0.88, 0.91, 0.9, 1),
                line_width=1,
                elevation=2,
                size_hint_y=None,
                height=dp(118),
            )
            caisse_float = MDFloatLayout()
            caisse_float.add_widget(
                MDIcon(
                    icon='cash-multiple',
                    theme_text_color='Custom',
                    text_color=(0.05, 0.55, 0.25, 1),
                    font_size='28sp',
                    pos_hint={'x': 0.04, 'top': 0.88},
                    size_hint=(None, None),
                    size=(dp(30), dp(30)),
                )
            )
            caisse_center = MDBoxLayout(
                orientation='vertical',
                adaptive_height=True,
                spacing=dp(4),
                padding=[dp(12), 0],
                pos_hint={'center_x': 0.5, 'center_y': 0.48},
            )
            caisse_center.add_widget(
                MDLabel(
                    text='TOTAL EN CAISSE',
                    halign='center',
                    font_style='Subtitle1',
                    bold=True,
                    theme_text_color='Primary',
                    adaptive_height=True,
                )
            )
            caisse_center.add_widget(
                MDLabel(
                    text='(Ventes payées + Versements - RC Comptoir - Dépenses)',
                    halign='center',
                    font_style='Caption',
                    theme_text_color='Hint',
                    adaptive_height=True,
                )
            )
            self.lbl_stat_caisse = MDLabel(
                text='0,00 DA',
                halign='center',
                font_style='H5',
                bold=True,
                theme_text_color='Custom',
                text_color=(0.1, 0.1, 0.1, 1),
                adaptive_height=True,
            )
            caisse_center.add_widget(self.lbl_stat_caisse)
            caisse_float.add_widget(caisse_center)
            caisse_card.add_widget(caisse_float)
            self.stats_card_container.add_widget(caisse_card)

            self.is_stats_expanded = False
            self.btn_stats_toggle = MDCard(
                orientation='horizontal',
                size_hint_y=None,
                height=dp(42),
                radius=[12],
                md_bg_color=(0.93, 0.95, 0.99, 1),
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
                text_color=self.theme_cls.primary_color,
                adaptive_width=True,
                pos_hint={'center_y': 0.5},
            )
            self.ico_stats_toggle = MDIcon(
                icon='chevron-down',
                theme_text_color='Custom',
                text_color=self.theme_cls.primary_color,
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
            stats_grid.add_widget(_stat_card('cart-outline', 'Ventes Nettes', 'lbl_stat_sales', (0.9, 0.95, 1, 1), (0, 0.35, 0.8, 1)))
            stats_grid.add_widget(_stat_card('cash-plus', 'Encaissements', 'lbl_stat_in', (0.92, 0.98, 0.92, 1), (0, 0.6, 0.2, 1)))
            stats_grid.add_widget(_stat_card('truck-outline', 'Achats Nets', 'lbl_stat_purchases', (1, 0.95, 0.86, 1), (0.9, 0.5, 0, 1)))
            stats_grid.add_widget(_stat_card('cash-minus', 'Décaissements', 'lbl_stat_out', (1, 0.91, 0.91, 1), (0.8, 0.1, 0.1, 1)))
            stats_grid.add_widget(_stat_card('chart-line-variant', 'Bénéfice Net', 'lbl_stat_profit', (0.94, 0.9, 0.99, 1), (0.5, 0.1, 0.7, 1)))
            stats_grid.add_widget(_stat_card('cash-remove', 'Dépenses', 'lbl_stat_expenses', (1, 0.94, 0.88, 1), (0.85, 0.35, 0, 1)))
            stats_grid.add_widget(_stat_card('account-cash', 'Créances Clients', 'lbl_stat_client_debts', (0.9, 1, 0.92, 1), (0, 0.6, 0.2, 1)))
            stats_grid.add_widget(_stat_card('account-credit-card', 'Dettes Fourniss.', 'lbl_stat_supplier_dues', (1, 0.88, 0.88, 1), (0.8, 0, 0, 1)))
            self.stats_more_container.add_widget(stats_grid)

            stock_card = _stat_card('package-variant', 'Stock (Coût)', 'lbl_stat_stock_val', (0.92, 0.97, 0.97, 1), (0, 0.5, 0.5, 1))
            stock_card.height = dp(92)
            self.stats_more_container.add_widget(stock_card)

        self.filter_dashboard_stats('auj')

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
        inact, act = (0.5, 0.5, 0.5, 1), self.theme_cls.primary_color
        for b in (self.btn_stat_today, self.btn_stat_yesterday, self.btn_stat_week, self.btn_stat_month, self.btn_stat_date):
            b.md_bg_color = inact
        today = datetime.now().date()
        start_d, end_d = today, today

        if filter_type == 'custom' and specific_start:
            start_d, end_d = specific_start, (specific_end or specific_start)
            self.btn_stat_date.md_bg_color = act
            self.btn_stat_date.text = start_d.strftime('%d/%m') if start_d == end_d else 'CAL'
            self.lbl_dashboard_date.text = f"Tableau de Bord ({start_d.strftime('%d/%m')} - {end_d.strftime('%d/%m')})" if start_d != end_d else f"Tableau de Bord ({start_d.strftime('%d/%m/%Y')})"
        else:
            self.btn_stat_date.text = 'CAL'
            if filter_type == 'auj':
                self.btn_stat_today.md_bg_color = act
                self.lbl_dashboard_date.text = "Tableau de Bord (Aujourd'hui)"
            elif filter_type == 'hier':
                start_d = end_d = today - timedelta(days=1)
                self.btn_stat_yesterday.md_bg_color = act
                self.lbl_dashboard_date.text = 'Tableau de Bord (Hier)'
            elif filter_type == 'semaine':
                start_d = today - timedelta(days=(today.weekday() + 2) % 7)
                self.btn_stat_week.md_bg_color = act
                self.lbl_dashboard_date.text = 'Tableau de Bord (Cette semaine)'
            elif filter_type == 'mois':
                start_d = today.replace(day=1)
                self.btn_stat_month.md_bg_color = act
                self.lbl_dashboard_date.text = 'Tableau de Bord (Ce mois)'

        self.stats_start_date = str(start_d)
        self.stats_end_date = str(end_d)
        self.check_and_load_stats()

    def check_and_load_stats(self):
        today_str = datetime.now().strftime('%Y-%m-%d')
        s_date = getattr(self, 'stats_start_date', today_str)
        e_date = getattr(self, 'stats_end_date', today_str)
        st = self.db.get_comprehensive_stats(s_date, e_date)
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
    def _build_products_screen(self):
        screen = MDScreen(name='products')
        root = MDBoxLayout(orientation='vertical')
        header = MDBoxLayout(orientation='vertical', adaptive_height=True, md_bg_color=(1, 1, 1, 1))
        self.prod_toolbar = MDTopAppBar(title='Produits', left_action_items=[['arrow-left', lambda x: self.go_back()]], elevation=2)
        header.add_widget(self.prod_toolbar)

        self.prod_search_layout = MDBoxLayout(orientation='horizontal', padding=[dp(10), dp(5), dp(10), 0], spacing=dp(5), size_hint_y=None, height=dp(55), md_bg_color=(1, 1, 1, 1))
        self.search_field = SmartTextField(hint_text='Rechercher...', mode='rectangle', icon_right='magnify', size_hint_y=None, height=dp(40), pos_hint={'center_y': 0.5}, size_hint_x=1)
        self.search_field.bind(text=self.filter_products)
        self.prod_search_layout.add_widget(self.search_field)

        self.btn_scan_prod = MDIconButton(icon='barcode-scan', theme_text_color='Custom', text_color=(0.2, 0.2, 0.2, 1), pos_hint={'center_y': 0.5}, icon_size='28sp', on_release=self.open_barcode_scanner)
        self.prod_search_layout.add_widget(self.btn_scan_prod)
        self.btn_add_prod = MDIconButton(icon='plus-circle', theme_text_color='Custom', text_color=(0, 0.7, 0, 1), pos_hint={'center_y': 0.5}, icon_size='36sp', on_release=lambda x: self.show_manage_product_dialog(None))
        self.prod_search_layout.add_widget(self.btn_add_prod)
        header.add_widget(self.prod_search_layout)
        header.add_widget(MDBoxLayout(size_hint_y=None, height=dp(5)))

        filter_box = MDBoxLayout(orientation='vertical', size_hint_y=None, height=dp(50), padding=[dp(10), 0, dp(10), dp(5)])
        family_card = MDCard(size_hint=(1, 1), radius=[5], md_bg_color=(0.96, 0.96, 0.96, 1), line_color=(0.6, 0.6, 0.6, 1), line_width=1, elevation=0, ripple_behavior=True, on_release=self.open_filter_menu)
        card_c = MDBoxLayout(orientation='horizontal', padding=[dp(15), 0])
        self.btn_main_family_filter = MDLabel(text=AppConstants.FAM_ALL, theme_text_color='Custom', text_color=(0.15, 0.15, 0.15, 1), font_style='Subtitle1', bold=True, halign='left', valign='center', font_name='ArabicFont')
        card_c.add_widget(self.btn_main_family_filter)
        card_c.add_widget(MDIcon(icon='menu-down', pos_hint={'center_y': 0.5}, theme_text_color='Secondary'))
        family_card.add_widget(card_c)
        filter_box.add_widget(family_card)
        header.add_widget(filter_box)
        header.add_widget(MDSeparator(height=dp(1), color=(0.9, 0.9, 0.9, 1)))
        root.add_widget(header)

        self.rv_products = ProductRecycleView()
        self.rv_products.size_hint_y = 1
        root.add_widget(self.rv_products)

        self.cart_bar = MDCard(size_hint_y=None, height=dp(60), padding=[dp(15), dp(5)], md_bg_color=self.theme_cls.primary_color, radius=[15, 15, 0, 0], ripple_behavior=True, on_release=self.open_cart_screen, elevation=4)
        cart_box = MDBoxLayout(orientation='horizontal')
        self.lbl_cart_count = MDLabel(text='PANIER (0)', theme_text_color='Custom', text_color=(1, 1, 1, 1), bold=True, halign='left', size_hint_x=0.5, font_style='Subtitle1')
        self.lbl_cart_total = MDLabel(text='0,00 DA', theme_text_color='Custom', text_color=(1, 1, 1, 1), bold=True, halign='right', font_style='H6', size_hint_x=0.5)
        cart_box.add_widget(self.lbl_cart_count)
        cart_box.add_widget(self.lbl_cart_total)
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

        self.prod_toolbar.title = AppConstants.MODE_TITLES.get(mode, 'Produits')
        self.theme_cls.primary_palette = AppConstants.MODE_COLORS.get(mode, 'Blue')

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
        else:
            _enter()

    def open_filter_menu(self, instance):
        families = self.db.get_families()
        full_list = [AppConstants.FAM_ALL] + [f for f in families if f != AppConstants.FAM_ALL]
        content = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True)
        scroll = MDScrollView(size_hint_y=None, height=dp(300))
        lst = MDBoxLayout(orientation='vertical', adaptive_height=True)
        for fam in full_list:
            lst.add_widget(
                MDFlatButton(
                    text=self.fix_text(fam),
                    font_name='ArabicFont',
                    theme_text_color='Custom',
                    text_color=(0.2, 0.2, 0.2, 1),
                    font_style='Subtitle1',
                    size_hint_x=1,
                    height=dp(48),
                    on_release=lambda x, f=fam: self.apply_filter(f),
                )
            )
            lst.add_widget(MDSeparator(height=dp(1), color=(0.9, 0.9, 0.9, 1)))
        scroll.add_widget(lst)
        content.add_widget(scroll)
        self.filter_dialog = MDDialog(title='Filtrer par famille', type='custom', content_cls=content, size_hint=(0.85, None))
        self.filter_dialog.open()

    def apply_filter(self, family_name: str):
        self.selected_family_filter = family_name
        self.btn_main_family_filter.text = self.fix_text(family_name)
        if getattr(self, 'filter_dialog', None):
            self.filter_dialog.dismiss()
        self.load_more_products(reset=True)

    def filter_products(self, instance, text):
        if self._search_event:
            self._search_event.cancel()
        self._search_event = Clock.schedule_once(lambda dt: self.load_more_products(reset=True), 0.25)

    def load_more_products(self, reset: bool = False):
        if self.is_loading_more and not reset:
            return
        if reset:
            self.current_page_offset = 0
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
                if is_virt:
                    price_fmt = 'Prix Libre'
                    price_col = [0, 0.4, 0.8, 1]
                    icon_name = 'pencil-plus'
                    icon_col = [0, 0.4, 0.8, 1]
                    stock_txt = 'Illimité'
                elif is_qty_only:
                    price_fmt = f'Tot : {FinancialEngine.format_qty(s_store + s_wh)}'
                    price_col = [0.2, 0.2, 0.8, 1]
                    stock_txt = 'Illimité' if is_unlim else f'Mag : {FinancialEngine.format_qty(s_store)} | Dép : {FinancialEngine.format_qty(s_wh)}'
                elif is_purchase:
                    p_val = FinancialEngine.to_decimal(p.get('purchase_price', 0))
                    price_fmt = FinancialEngine.format_money(p_val)
                    price_col = [0.9, 0.5, 0, 1]
                    stock_txt = 'Illimité' if is_unlim else (f'Mag : {FinancialEngine.format_qty(s_store)} | Dép : {FinancialEngine.format_qty(s_wh)}' if s_wh != 0 else f'Qté : {FinancialEngine.format_qty(s_store)}')
                else:
                    disp_price, has_promo = self._compute_product_sale_price(p, check_promo=show_promo, today_date=today_d)
                    price_fmt = FinancialEngine.format_money(disp_price)
                    if has_promo:
                        price_col = [0, 0.2, 0.8, 1]
                        icon_name = 'sale'
                        icon_col = [0.9, 0.1, 0.1, 1]
                    else:
                        price_col = [0, 0.6, 0.2, 1]
                    if is_unlim:
                        stock_txt = 'Illimité'
                    elif s_wh != 0:
                        stock_txt = f'Qté : {FinancialEngine.format_qty(s_store)} | Dép : {FinancialEngine.format_qty(s_wh)}'
                    else:
                        stock_txt = f'Qté : {FinancialEngine.format_qty(s_store)}'
                    if not has_promo and not is_unlim and (s_store + s_wh) <= 0:
                        icon_name = 'package-variant-closed'
                        icon_col = [0.8, 0, 0, 1]

                rv_data.append({
                    'text_name': self.fix_text(p.get('name', '')),
                    'text_price': price_fmt,
                    'text_stock': stock_txt,
                    'icon_name': icon_name,
                    'icon_color': icon_col,
                    'price_color': price_col,
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
        if self.rv_products:
            if reset:
                self.rv_products.data = new_data
                self.rv_products.scroll_y = 1.0
                self.current_page_offset = db_count
            else:
                self.rv_products.data.extend(new_data)
                self.current_page_offset += db_count
            self.rv_products.refresh_from_data()
            self.rv_products.loading_lock = False
        self.is_loading_more = False

    def _finish_products_loading(self):
        self.is_loading_more = False
        if self.rv_products:
            self.rv_products.loading_lock = False

    def show_zoomed_image(self, image_path: str, title: str = 'Image'):
        if not image_path or not os.path.exists(image_path):
            self.notify('Image non disponible', 'error')
            return
        content = MDBoxLayout(orientation='vertical', padding=0, spacing=0, size_hint_y=None, height=dp(400))
        content.add_widget(FitImage(source=image_path, radius=[10, 10, 0, 0], size_hint=(1, 1), mipmap=True))
        content.add_widget(MDRaisedButton(text='FERMER', md_bg_color=(0.2, 0.2, 0.2, 1), text_color=(1, 1, 1, 1), size_hint_x=1, on_release=lambda x: self.zoom_dialog.dismiss()))
        self.zoom_dialog = MDDialog(title=title, type='custom', content_cls=content, size_hint=(0.9, None))
        self.zoom_dialog.open()

    # --- Gestion des Images Produit ---
    def open_image_selector(self, instance):
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
            btns.add_widget(MDFlatButton(text='ANNULER', on_release=lambda x: self.img_picker_dialog.dismiss()))
            btns.add_widget(MDRaisedButton(text='SÉLECTIONNER', md_bg_color=(0.1, 0.4, 0.8, 1), on_release=lambda x: self.confirm_image_selection(self.img_file_chooser.selection)))
            content.add_widget(btns)
            self.img_picker_dialog = MDDialog(title='Choisir une image', type='custom', content_cls=content, size_hint=(0.95, 0.9))
            self.img_picker_dialog.open()

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
                Clock.schedule_once(lambda dt: self.confirm_image_selection([temp_path]), 0)
        except Exception as exc:
            Clock.schedule_once(lambda dt, e=exc: self.notify(f'Erreur image : {e}', 'error'), 0)

    def confirm_image_selection(self, selection):
        if getattr(self, 'img_picker_dialog', None):
            self.img_picker_dialog.dismiss()
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
                if img.mode in ('RGBA', 'P'):
                    img = img.convert('RGB')
                img.thumbnail((1024, 1024), Image.Resampling.LANCZOS)
                img.save(dest_path, quality=85, optimize=True)
            return dest_path
        except Exception:
            shutil.copyfile(source_path, dest_path)
            return dest_path

        # --- Familles (Dialogues) ---
    def open_family_selector_dialog(self):
        families = [f for f in self.db.get_families() if f != AppConstants.FAM_ALL]
        full_list = [AppConstants.FAM_ALL] + families
        content = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True)
        scroll = MDScrollView(size_hint_y=None, height=dp(300))
        lst = MDList()
        for fam in full_list:
            item = OneLineAvatarIconListItem(text=self.fix_text(fam), on_release=lambda x, f=fam: self._select_family_and_close(f))
            if fam != AppConstants.FAM_ALL:
                del_ic = IconRightWidget(icon='trash-can-outline', theme_text_color='Custom', text_color=(0.8, 0, 0, 1))
                del_ic.bind(on_release=lambda x, f=fam: self.show_delete_family_confirmation(f))
                item.add_widget(del_ic)
            lst.add_widget(item)
        scroll.add_widget(lst)
        content.add_widget(scroll)
        self.family_dialog = MDDialog(title='Sélectionner une famille', type='custom', content_cls=content, size_hint=(0.85, None))
        self.family_dialog.open()

    def _select_family_and_close(self, family_name: str):
        self.selected_product_family_raw = family_name
        if hasattr(self, 'btn_select_family'):
            self.btn_select_family.text = self.fix_text(family_name)
        if getattr(self, 'family_dialog', None):
            self.family_dialog.dismiss()

    def show_add_family_dialog(self):
        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, height=dp(100), padding=dp(10))
        field = SmartTextField(hint_text='Nom de la nouvelle famille')
        content.add_widget(field)

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

        self.add_fam_dialog = MDDialog(
            title='Nouvelle Famille',
            type='custom',
            content_cls=content,
            buttons=[
                MDFlatButton(text='ANNULER', on_release=lambda x: self.add_fam_dialog.dismiss()),
                MDRaisedButton(text='AJOUTER', on_release=_save),
            ],
        )
        self.add_fam_dialog.open()

    def show_delete_family_confirmation(self, family_name: str):
        def _do_del(x):
            self.confirm_del_fam_dialog.dismiss()
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

        self.confirm_del_fam_dialog = MDDialog(
            title='Confirmation',
            text=f"Voulez-vous vraiment supprimer la famille '{family_name}' ?",
            buttons=[
                MDFlatButton(text='ANNULER', on_release=lambda x: self.confirm_del_fam_dialog.dismiss()),
                MDRaisedButton(text='SUPPRIMER', md_bg_color=(0.8, 0, 0, 1), on_release=_do_del),
            ],
        )
        self.confirm_del_fam_dialog.open()

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
        val_tva = _fmt_p(product.get('tva', 0)) if is_edit else ''

        val_is_promo = bool(int(product.get('is_promo_active', 0) or 0) == 1) if is_edit else False
        self.promo_type_selected = product.get('promo_type', 'fixed') if is_edit else 'fixed'
        val_promo_val = _fmt_p(product.get('promo_value', 0)) if is_edit else ''
        val_promo_exp = str(product.get('promo_expiry', '') or '') if is_edit else ''

        scroll = MDScrollView(size_hint_y=None, height=dp(600))
        main_box = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(14), padding=[dp(10), dp(10), dp(10), dp(20)])

        card_info = MDCard(orientation='vertical', radius=[12], padding=dp(15), spacing=dp(10), elevation=1, adaptive_height=True, md_bg_color=(0.98, 0.98, 0.98, 1))
        card_info.add_widget(MDLabel(text='Informations Produit', bold=True, theme_text_color='Primary', font_style='Subtitle1'))
        self.field_num = MDTextField(text=val_num, hint_text='N° Produit', icon_right='pound')
        card_info.add_widget(self.field_num)

        box_bar = MDBoxLayout(orientation='horizontal', spacing=dp(5), adaptive_height=True)
        self.field_bar = MDTextField(text=val_barcode, hint_text='Code-Barres', size_hint_x=1, icon_right='barcode')
        box_bar.add_widget(self.field_bar)
        box_bar.add_widget(MDIconButton(icon='barcode-scan', on_release=lambda x: [setattr(self, 'target_scan_field', self.field_bar), self.open_barcode_scanner(None)]))
        box_bar.add_widget(MDIconButton(icon='refresh', theme_text_color='Custom', text_color=self.theme_cls.primary_color, on_release=lambda x: setattr(self.field_bar, 'text', self._generate_ean13())))
        card_info.add_widget(box_bar)

        self.field_name = SmartTextField(text=val_name, hint_text='Désignation *', required=True, icon_right='tag-text-outline')
        self.field_reference = SmartTextField(text=val_ref, hint_text='Référence', icon_right='text')
        card_info.add_widget(self.field_name)
        card_info.add_widget(self.field_reference)

        card_info.add_widget(MDLabel(text='Famille :', font_style='Caption', theme_text_color='Secondary'))
        fam_card = MDCard(size_hint_y=None, height=dp(48), radius=[6], md_bg_color=(0.95, 0.95, 0.95, 1), elevation=0)
        fam_float = MDFloatLayout()
        self.btn_select_family = MDFlatButton(text=self.fix_text(self.selected_product_family_raw), font_name='ArabicFont', size_hint=(1, 1), pos_hint={'center_x': 0.5, 'center_y': 0.5}, on_release=lambda x: self.open_family_selector_dialog())
        fam_float.add_widget(self.btn_select_family)
        fam_float.add_widget(MDIconButton(icon='plus-circle', theme_text_color='Custom', text_color=(0, 0.65, 0.2, 1), pos_hint={'right': 1, 'center_y': 0.5}, on_release=lambda x: self.show_add_family_dialog()))
        fam_card.add_widget(fam_float)
        card_info.add_widget(fam_card)

        img_status = os.path.basename(current_image) if (current_image and os.path.exists(current_image)) else 'Aucune'
        box_img = MDBoxLayout(orientation='horizontal', spacing=dp(8), adaptive_height=True, padding=[0, dp(6), 0, 0])
        self.lbl_image_status = MDLabel(text=f'Image : {img_status}', font_style='Caption', size_hint_x=0.5, theme_text_color='Secondary', shorten=True)
        box_img.add_widget(self.lbl_image_status)
        box_img.add_widget(MDIconButton(icon='trash-can-outline', theme_text_color='Custom', text_color=(0.8, 0, 0, 1), on_release=lambda x: [setattr(self, 'temp_selected_image_path', None), setattr(self, 'remove_image_order', True), setattr(self.lbl_image_status, 'text', 'Image : Aucune')]))
        box_img.add_widget(MDRaisedButton(text='Choisir', size_hint_x=0.3, md_bg_color=(0.2, 0.5, 0.8, 1), on_release=self.open_image_selector))
        card_info.add_widget(box_img)
        main_box.add_widget(card_info)

        card_stock = MDCard(orientation='vertical', radius=[12], padding=dp(15), spacing=dp(10), elevation=1, adaptive_height=True, md_bg_color=(0.96, 0.99, 0.96, 1))
        card_stock.add_widget(MDLabel(text='Stock Initial (Magasin)', bold=True, theme_text_color='Primary', font_style='Subtitle1'))
        row_stock = MDBoxLayout(orientation='horizontal', spacing=dp(10), adaptive_height=True)
        self.chk_unlimited = MDCheckbox(active=is_unlimited, size_hint=(None, None), size=(dp(40), dp(40)))
        lbl_unlim = MDLabel(text='Illimité\n(Verrouillé)' if has_movements else 'Illimité', font_style='Body2', size_hint_x=0.45)
        self.field_stock = MDTextField(text=val_stock, hint_text='Quantité', input_filter='float', size_hint_x=0.55)
        if has_movements:
            self.chk_unlimited.disabled = True
            self.field_stock.disabled = True
            self.field_stock.helper_text = 'Verrouillé (mouvements existants)'
            self.field_stock.helper_text_mode = 'persistent'
        else:
            self.chk_unlimited.bind(active=lambda c, v: setattr(self.field_stock, 'disabled', bool(v)))
            self.field_stock.disabled = is_unlimited
        row_stock.add_widget(self.chk_unlimited)
        row_stock.add_widget(lbl_unlim)
        row_stock.add_widget(self.field_stock)
        card_stock.add_widget(row_stock)
        main_box.add_widget(card_stock)

        card_price = MDCard(orientation='vertical', radius=[12], padding=dp(15), spacing=dp(10), elevation=1, adaptive_height=True, md_bg_color=(0.96, 0.96, 0.99, 1))
        card_price.add_widget(MDLabel(text='Tarification (DA)', bold=True, theme_text_color='Primary', font_style='Subtitle1'))
        self.field_cost = MDTextField(text=val_cost, hint_text="Prix d'Achat (Coût / PMP)", input_filter='float')
        if has_movements:
            self.field_cost.readonly = True
            self.field_cost.helper_text = 'PMP calculé automatiquement'
            self.field_cost.helper_text_mode = 'persistent'
        self.field_p1 = MDTextField(text=val_p1, hint_text='Prix Détail', input_filter='float')
        self.field_p2 = MDTextField(text=val_p2, hint_text='Prix Demi-Gros', input_filter='float')
        self.field_p3 = MDTextField(text=val_p3, hint_text='Prix Gros', input_filter='float')
        self.field_tva = MDTextField(text=val_tva, hint_text='TVA % (Factures)', input_filter='float')
        for w in (self.field_cost, self.field_p1, self.field_p2, self.field_p3, self.field_tva):
            card_price.add_widget(w)
        main_box.add_widget(card_price)

        card_promo = MDCard(orientation='vertical', radius=[12], padding=dp(15), spacing=dp(10), elevation=1, adaptive_height=True, md_bg_color=(1, 0.95, 0.95, 1))
        h_promo = MDBoxLayout(orientation='horizontal', spacing=dp(10), adaptive_height=True)
        h_promo.add_widget(MDIcon(icon='sale', theme_text_color='Custom', text_color=(0.8, 0, 0, 1)))
        h_promo.add_widget(MDLabel(text='Promotion', bold=True, theme_text_color='Custom', text_color=(0.8, 0, 0, 1)))
        self.chk_promo_active = MDCheckbox(active=val_is_promo, size_hint=(None, None), size=(dp(40), dp(40)))
        h_promo.add_widget(self.chk_promo_active)
        card_promo.add_widget(h_promo)

        type_box = MDBoxLayout(orientation='horizontal', spacing=dp(10), adaptive_height=True)
        self.btn_promo_fixed = MDRaisedButton(text='Prix Fixe', size_hint_x=0.5, elevation=0)
        self.btn_promo_percent = MDRaisedButton(text='Pourcentage %', size_hint_x=0.5, elevation=0)
        self.field_promo_val = MDTextField(text=val_promo_val, hint_text='Valeur', input_filter='float')

        def _upd_promo_btns():
            is_fx = self.promo_type_selected == 'fixed'
            self.btn_promo_fixed.md_bg_color = (0.8, 0, 0, 1) if is_fx else (0.9, 0.9, 0.9, 1)
            self.btn_promo_fixed.text_color = (1, 1, 1, 1) if is_fx else (0, 0, 0, 1)
            self.btn_promo_percent.md_bg_color = (0.9, 0.9, 0.9, 1) if is_fx else (0.8, 0, 0, 1)
            self.btn_promo_percent.text_color = (0, 0, 0, 1) if is_fx else (1, 1, 1, 1)
            self.field_promo_val.hint_text = 'Nouveau Prix (DA)' if is_fx else 'Remise (%)'

        self.btn_promo_fixed.bind(on_release=lambda x: [setattr(self, 'promo_type_selected', 'fixed'), _upd_promo_btns()])
        self.btn_promo_percent.bind(on_release=lambda x: [setattr(self, 'promo_type_selected', 'percent'), _upd_promo_btns()])
        type_box.add_widget(self.btn_promo_fixed)
        type_box.add_widget(self.btn_promo_percent)
        card_promo.add_widget(type_box)
        card_promo.add_widget(self.field_promo_val)

        box_date = MDBoxLayout(orientation='horizontal', spacing=dp(5), adaptive_height=True)
        self.btn_promo_date = MDRaisedButton(text=val_promo_exp if val_promo_exp else 'Date Fin (Optionnel)', md_bg_color=(0.3, 0.3, 0.3, 1), size_hint_x=0.85)

        def _open_promo_picker(x):
            dp_dlg = MDDatePicker()
            dp_dlg.bind(on_save=lambda inst, val, dr: setattr(self.btn_promo_date, 'text', str(val)))
            dp_dlg.open()

        self.btn_promo_date.bind(on_release=_open_promo_picker)
        box_date.add_widget(self.btn_promo_date)
        box_date.add_widget(MDIconButton(icon='calendar-remove', theme_text_color='Custom', text_color=(0.8, 0.2, 0.2, 1), size_hint_x=0.15, on_release=lambda x: setattr(self.btn_promo_date, 'text', 'Date Fin (Optionnel)')))
        card_promo.add_widget(box_date)
        _upd_promo_btns()
        main_box.add_widget(card_promo)

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
                    'tva': FinancialEngine.to_decimal(self.field_tva.text or '0'),
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

            def _confirm(y):
                if getattr(self, 'conf_diag', None):
                    self.conf_diag.dismiss()
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

            self.conf_diag = MDDialog(
                title='Confirmation',
                text='Voulez-vous vraiment supprimer ce produit ?',
                buttons=[
                    MDFlatButton(text='NON', on_release=lambda z: self.conf_diag.dismiss()),
                    MDRaisedButton(text='OUI, SUPPRIMER', md_bg_color=(0.8, 0, 0, 1), on_release=_confirm),
                ],
            )
            self.conf_diag.open()

        footer = MDBoxLayout(orientation='vertical', spacing=dp(10), adaptive_height=True, padding=[0, dp(16), 0, 0])
        footer.add_widget(MDRaisedButton(text='ENREGISTRER', md_bg_color=(0, 0.7, 0.2, 1), text_color=(1, 1, 1, 1), size_hint_x=1, height=dp(50), on_release=_save_product))
        if is_edit:
            footer.add_widget(MDRaisedButton(text='SUPPRIMER', md_bg_color=(0.8, 0, 0, 1), text_color=(1, 1, 1, 1), size_hint_x=1, height=dp(50), on_release=_delete_product))
        footer.add_widget(MDRaisedButton(text='FERMER', md_bg_color=(0.9, 0.9, 0.9, 1), text_color=(0.2, 0.2, 0.2, 1), size_hint_x=1, height=dp(50), on_release=lambda x: self.dialog.dismiss()))
        main_box.add_widget(footer)
        scroll.add_widget(main_box)

        self.dialog = MDDialog(title='Fiche Produit' if is_edit else 'Nouveau Produit', type='custom', content_cls=scroll, buttons=[], size_hint=(0.95, 0.88))
        self.dialog.open()
        Clock.schedule_once(lambda dt: setattr(self.field_name, 'focus', True), 0.3)

    # --------------------------------------------------------------------------
    # 7.5 CONTRÔLE DE STOCK, PANIER & RÈGLES MÉTIER DESKTOP
    # --------------------------------------------------------------------------
    def _is_negative_stock_enforced(self) -> bool:
        doc_type = AppConstants.get_doc_type_for_mode(self.current_mode)
        if doc_type == AppConstants.DOC_TR:
            return True
        if self.db.get_bool_setting('allow_negative_stock', default=False):
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
        self.stock_error_dialog = MDDialog(
            title='Stock Insuffisant',
            text=message,
            buttons=[MDRaisedButton(text='COMPRIS', md_bg_color=(0.8, 0, 0, 1), text_color=(1, 1, 1, 1), on_release=lambda x: self.stock_error_dialog.dismiss())],
            size_hint=(0.88, None),
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

        ok_stk, msg_stk = self._check_stock_availability(product, Decimal('1'))
        if not ok_stk and self.sm.current == 'cart':
            self.play_sound('error')
            self._show_stock_error_dialog(msg_stk)
            return

        curr_price, has_promo, price_found_in_history = self._resolve_product_unit_price(product, doc_type)

        self.active_input_target = 'qty'
        self.input_reset_mode = True

        def _upd_colors():
            act, inact = (0.9, 1, 0.9, 1), (0.95, 0.95, 0.95, 1)
            if hasattr(self, 'qty_card'):
                self.qty_card.md_bg_color = act if self.active_input_target == 'qty' else inact
                self.qty_card.elevation = 3 if self.active_input_target == 'qty' else 0
            if hasattr(self, 'price_card'):
                self.price_card.md_bg_color = act if self.active_input_target == 'price' else inact
                self.price_card.elevation = 3 if self.active_input_target == 'price' else 0

        header_box = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing='4dp')
        header_box.add_widget(MDLabel(text=self.fix_text(product.get('name')), halign='center', bold=True, font_style='Subtitle1', theme_text_color='Primary', adaptive_height=True))
        if has_promo:
            header_box.add_widget(MDLabel(text='PROMOTION ACTIVÉE', halign='center', theme_text_color='Custom', text_color=(0.85, 0, 0, 1), font_style='Caption', bold=True, adaptive_height=True))
        elif price_found_in_history:
            header_box.add_widget(MDLabel(text='Dernier Prix Appliqué (Historique)', halign='center', theme_text_color='Custom', text_color=(0, 0.45, 0.75, 1), font_style='Caption', bold=True, adaptive_height=True))

        content = MDBoxLayout(orientation='vertical', spacing='8dp', size_hint_y=None, height=dp(410 if is_qty_only else 500), padding=[0, '5dp', 0, 0])
        content.add_widget(header_box)

        if not is_qty_only:
            self.price_card = MDCard(size_hint_y=None, height='70dp', radius=[10], padding=[10, 0, 10, 0], elevation=0)
            self.price_field = NoMenuTextField(
                text=FinancialEngine.format_qty(curr_price),
                hint_text='Dernier Prix (Historique)' if price_found_in_history else 'Prix Unitaire (DA)',
                font_size='26sp',
                halign='center',
                mode='line',
                readonly=True,
                line_color_normal=(0, 0, 0, 0),
                line_color_focus=(0, 0, 0, 0),
                pos_hint={'center_y': 0.5},
            )
            self.price_field.theme_text_color = 'Custom'
            price_txt_col = (0, 0.45, 0.75, 1) if price_found_in_history else (0, 0, 0, 1)
            self.price_field.text_color_normal = price_txt_col
            self.price_field.text_color_focus = price_txt_col

            def _on_price_touch(inst, touch):
                if inst.collide_point(*touch.pos):
                    if self.active_input_target != 'price':
                        self.input_reset_mode = True
                    self.active_input_target = 'price'
                    _upd_colors()
                    return True
                return False

            self.price_field.bind(on_touch_down=_on_price_touch)
            self.price_card.add_widget(self.price_field)
            p_row = MDBoxLayout(size_hint_y=None, height='75dp', padding=[60, 0, 60, 0])
            p_row.add_widget(self.price_card)
            content.add_widget(p_row)
        elif hasattr(self, 'price_field'):
            del self.price_field
            if hasattr(self, 'price_card'):
                del self.price_card

        qty_row = MDBoxLayout(orientation='horizontal', spacing='10dp', size_hint_y=None, height='65dp', padding=[40, 0])
        btn_minus = MDIconButton(icon='minus', theme_text_color='Custom', text_color=(1, 1, 1, 1), md_bg_color=(0.9, 0.3, 0.3, 1), pos_hint={'center_y': 0.5})
        self.qty_card = MDCard(size_hint_x=1, size_hint_y=None, height='60dp', radius=[10], padding=[10, 0, 10, 0], elevation=0, pos_hint={'center_y': 0.5})
        self.qty_field = NoMenuTextField(
            text='1',
            hint_text='Qté',
            font_size='28sp',
            halign='center',
            readonly=True,
            mode='line',
            line_color_normal=(0, 0, 0, 0),
            line_color_focus=(0, 0, 0, 0),
            pos_hint={'center_y': 0.5},
        )
        self.qty_field.theme_text_color = 'Custom'
        self.qty_field.text_color_normal = (0, 0, 0, 1)
        self.qty_field.text_color_focus = (0, 0, 0, 1)

        def _on_qty_touch(inst, touch):
            if inst.collide_point(*touch.pos):
                if self.active_input_target != 'qty':
                    self.input_reset_mode = True
                self.active_input_target = 'qty'
                _upd_colors()
                return True
            return False

        self.qty_field.bind(on_touch_down=_on_qty_touch)
        self.qty_card.add_widget(self.qty_field)
        btn_plus = MDIconButton(icon='plus', theme_text_color='Custom', text_color=(1, 1, 1, 1), md_bg_color=(0.2, 0.7, 0.2, 1), pos_hint={'center_y': 0.5})
        qty_row.add_widget(btn_minus)
        qty_row.add_widget(self.qty_card)
        qty_row.add_widget(btn_plus)
        content.add_widget(qty_row)

        self.btn_add = MDRaisedButton(text='AJOUTER', md_bg_color=(0, 0.7, 0.2, 1), text_color=(1, 1, 1, 1), size_hint_x=0.7, size_hint_y=1, font_size='18sp', elevation=3)

        def _update_btn_txt():
            if is_qty_only:
                self.btn_add.text = 'AJOUTER'
                return
            q = FinancialEngine.to_decimal(self.qty_field.text or '0')
            p = FinancialEngine.to_decimal(self.price_field.text or '0')
            calc = FinancialEngine.calculate_document_totals([{'id': product['id'], 'qty': q, 'price': p, 'tva': product.get('tva', 0)}], doc_type=doc_type)
            self.btn_add.text = f"AJOUTER\n{FinancialEngine.format_money(calc['final_total'])}"

        def _confirm_add(x):
            q_val = FinancialEngine.to_decimal(self.qty_field.text or '0')
            p_val = FinancialEngine.to_decimal(self.price_field.text or '0') if not is_qty_only else Decimal('0.00')
            ok, err = FinancialEngine.validate_item_bounds(q_val, p_val, product.get('tva', 0))
            if not ok:
                self.notify(err, 'error')
                return

            cost_val = FinancialEngine.to_decimal(product.get('purchase_price', 0))
            if is_sale and not AppConstants.is_default_product(product) and self.db.get_bool_setting('warn_price_below_cost', default=True):
                if cost_val > 0 and p_val < cost_val:
                    def _proceed_low(y):
                        self.warn_cost_dlg.dismiss()
                        self._finalize_add_to_cart(product, q_val, p_val, doc_type)

                    self.warn_cost_dlg = MDDialog(
                        title='Attention : Prix bas',
                        text=f"Le prix de vente ({FinancialEngine.format_money(p_val)}) est inférieur au coût d'achat ({FinancialEngine.format_money(cost_val)}).\n\nContinuer quand même ?",
                        buttons=[
                            MDFlatButton(text='ANNULER', on_release=lambda y: self.warn_cost_dlg.dismiss()),
                            MDRaisedButton(text='OUI, CONTINUER', md_bg_color=(0.9, 0.5, 0, 1), on_release=_proceed_low),
                        ],
                    )
                    self.warn_cost_dlg.open()
                    return
            self._finalize_add_to_cart(product, q_val, p_val, doc_type)

        self.btn_add.bind(on_release=_confirm_add)

        def _delta_q(d: int):
            v = max(Decimal('1'), FinancialEngine.to_decimal(self.qty_field.text or '0') + Decimal(d))
            self.qty_field.text = FinancialEngine.format_qty(v)
            _update_btn_txt()

        btn_plus.bind(on_release=lambda x: _delta_q(1))
        btn_minus.bind(on_release=lambda x: _delta_q(-1))

        def _active_f():
            if is_qty_only:
                return self.qty_field
            return self.price_field if self.active_input_target == 'price' else self.qty_field

        def _digit(d: str):
            fld = _active_f()
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
            self.input_reset_mode = False
            if fld.text:
                fld.text = fld.text[:-1]
            _update_btn_txt()

        grid = MDGridLayout(cols=3, spacing='8dp', size_hint_y=1, padding=[20, 0])
        for key in ('7', '8', '9', '4', '5', '6', '1', '2', '3', '.', '0', 'DEL'):
            if key == 'DEL':
                grid.add_widget(MDIconButton(icon='backspace-outline', theme_text_color='Custom', text_color=(0, 0, 0, 1), md_bg_color=(0.8, 0.8, 0.8, 1), size_hint=(1, 1), on_release=_bkspc))
            else:
                grid.add_widget(MDRaisedButton(text=key, md_bg_color=(0.95, 0.95, 0.95, 1), theme_text_color='Custom', text_color=(0, 0, 0, 1), font_size='22sp', size_hint=(1, 1), elevation=1, on_release=lambda x, k=key: _digit(k)))
        content.add_widget(grid)
        content.add_widget(MDLabel(text='', size_hint_y=None, height='8dp'))

        btns = MDBoxLayout(orientation='horizontal', spacing='10dp', size_hint_y=None, height='60dp')
        btns.add_widget(MDFlatButton(text='ANNULER', theme_text_color='Custom', text_color=(0.5, 0.5, 0.5, 1), size_hint_x=0.3, on_release=lambda x: self.dialog.dismiss()))
        btns.add_widget(self.btn_add)
        content.add_widget(btns)

        _upd_colors()
        _update_btn_txt()
        self.dialog = MDDialog(title='', type='custom', content_cls=content, buttons=[], size_hint=(0.85, None))
        self.dialog.open()

    def _finalize_add_to_cart(self, product: dict, qty: Decimal, price: Decimal, doc_type: str):
        ok_stk, msg_stk = self._check_stock_availability(product, qty)
        if not ok_stk:
            self.play_sound('error')
            self._show_stock_error_dialog(msg_stk)
            return

        tva_val = FinancialEngine.quantize_money(product.get('tva', 0)) if doc_type in AppConstants.INVOICE_TAX_DOC_TYPES else Decimal('0.00')
        if AppConstants.is_default_product(product):
            cnt = sum(1 for it in self.cart if AppConstants.is_default_product(it))
            new_name = f'{AppConstants.DEFAULT_PRODUCT_NAME} {cnt + 1}'
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
        if item in self.cart:
            self.cart.remove(item)
        self.refresh_cart_screen_items()
        self.update_cart_button()

    def recalculate_cart_prices(self):
        if not self.cart or not self.selected_entity:
            return
        doc_type = AppConstants.get_doc_type_for_mode(self.current_mode)
        if doc_type in AppConstants.QTY_ONLY_DOC_TYPES:
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
        screen = MDScreen(name='cart')
        layout = MDBoxLayout(orientation='vertical')
        self.cart_toolbar = MDTopAppBar(title='Panier', left_action_items=[['arrow-left', lambda x: self.back_to_products()]])
        layout.add_widget(self.cart_toolbar)

        selectors = MDCard(orientation='horizontal', size_hint_y=None, height=dp(70), padding=dp(10), radius=0, md_bg_color=(0.95, 0.95, 0.95, 1))
        self.btn_ent_screen = MDFillRoundFlatButton(text='Client', size_hint_x=0.45, on_release=self.handle_entity_button_click)
        self.btn_loc_screen = MDFillRoundFlatIconButton(text='Magasin', icon='store', size_hint_x=0.45, on_release=self.toggle_location)
        selectors.add_widget(self.btn_ent_screen)
        selectors.add_widget(MDBoxLayout(size_hint_x=0.1))
        selectors.add_widget(self.btn_loc_screen)
        layout.add_widget(selectors)

        self.rv_cart = CartRecycleView()
        layout.add_widget(self.rv_cart)

        self.footer_card = MDCard(orientation='vertical', size_hint_y=None, height=dp(150), padding=dp(15), spacing=dp(10), radius=[20, 20, 0, 0], elevation=4)
        total_row = MDBoxLayout(orientation='horizontal')
        self.lbl_total_title = MDLabel(text='TOTAL :', bold=True, font_style='Subtitle1')
        self.lbl_cart_screen_total = MDLabel(text='0,00 DA', halign='right', font_style='H5', bold=True, theme_text_color='Primary')
        total_row.add_widget(self.lbl_total_title)
        total_row.add_widget(self.lbl_cart_screen_total)
        self.btn_validate_cart = MDFillRoundFlatButton(text='VALIDER LA COMMANDE', size_hint_x=1, height=dp(55), md_bg_color=(0, 0.7, 0.2, 1), on_release=self.open_payment_dialog)
        self.footer_card.add_widget(total_row)
        self.footer_card.add_widget(self.btn_validate_cart)
        layout.add_widget(self.footer_card)
        screen.add_widget(layout)
        return screen

    def open_cart_screen(self, x=None):
        if not self.cart:
            self.dialog = MDDialog(title='Panier vide', text='Veuillez ajouter au moins un produit.', buttons=[MDFlatButton(text='OK', on_release=lambda x: self.dialog.dismiss())])
            self.dialog.open()
            return
        if self.current_mode != 'transfer' and self.selected_entity is None:
            self.show_entity_selection_dialog(None, next_action=lambda: self.open_cart_screen(None))
            return
        self.refresh_cart_screen_items()
        self.sm.transition.direction = 'left'
        self.sm.current = 'cart'

    def back_to_products(self):
        self.sm.transition.direction = 'right'
        self.sm.current = 'products'

    def handle_entity_button_click(self, instance):
        if self.current_mode == 'transfer':
            self.toggle_location(instance)
        else:
            self.show_entity_selection_dialog(instance)

    def toggle_location(self, x=None):
        self.selected_location = AppConstants.LOC_WAREHOUSE if self.selected_location == AppConstants.LOC_STORE else AppConstants.LOC_STORE
        self.update_location_display()
        if self.current_mode == 'transfer' and hasattr(self, 'btn_ent_screen'):
            src = AppConstants.LOC_LABELS[self.selected_location]
            dst = AppConstants.LOC_LABELS[AppConstants.LOC_WAREHOUSE if self.selected_location == AppConstants.LOC_STORE else AppConstants.LOC_STORE]
            self.btn_ent_screen.text = f'{src}  >>>  {dst}'

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
        is_qty_only = doc_type in AppConstants.QTY_ONLY_DOC_TYPES
        is_invoice = doc_type in AppConstants.INVOICE_TAX_DOC_TYPES
        hide_loc = is_transfer or doc_type in (AppConstants.DOC_FP, AppConstants.DOC_DP)

        ent_id = self.selected_entity.get('id') if self.selected_entity else None
        is_comptoir = AppConstants.is_default_entity(ent_id, AppConstants.get_entity_type(doc_type))
        calc = FinancialEngine.calculate_document_totals(
            self.cart,
            doc_type=doc_type,
            global_discount=self.editing_global_discount,
            payment_method=self.editing_payment_method,
            is_comptoir=is_comptoir,
        )

        self.cart_toolbar.title = f'Panier ({len(self.cart)})'
        self.update_location_display()

        if is_qty_only:
            self.lbl_total_title.text = 'Résumé :'
            self.lbl_cart_screen_total.text = ''
            self.lbl_cart_screen_total.opacity = 0
        else:
            self.lbl_total_title.text = 'TOTAL (+Timbre) :' if calc['stamp_duty'] > 0 else 'TOTAL :'
            self.lbl_cart_screen_total.text = FinancialEngine.format_money(calc['final_total'])
            self.lbl_cart_screen_total.opacity = 1

        if is_transfer:
            src = AppConstants.LOC_LABELS[self.selected_location]
            dst = AppConstants.LOC_LABELS[AppConstants.LOC_WAREHOUSE if self.selected_location == AppConstants.LOC_STORE else AppConstants.LOC_STORE]
            self.btn_ent_screen.text = f'{src}  >>>  {dst}'
            self.btn_ent_screen.md_bg_color = (0.5, 0, 0.5, 1)
            self.btn_validate_cart.text = 'VALIDER LE TRANSFERT'
            self.btn_validate_cart.md_bg_color = (0.5, 0, 0.5, 1)
        else:
            if self.selected_entity:
                self.btn_ent_screen.text = self.fix_text(self.selected_entity.get('name', 'Tiers'))[:15]
            vis = AppConstants.DOC_VISUALS.get(doc_type, {'color': (0, 0.6, 0.6, 1)})
            self.btn_ent_screen.md_bg_color = vis['color']
            self.btn_validate_cart.text = 'VALIDER LA COMMANDE'
            self.btn_validate_cart.md_bg_color = (0, 0.7, 0.2, 1)

        if hide_loc:
            self.btn_ent_screen.size_hint_x = 0.95
            self.btn_loc_screen.opacity = 0
            self.btn_loc_screen.disabled = True
            self.btn_loc_screen.size_hint_x = 0
            self.btn_loc_screen.width = 0
        else:
            self.btn_ent_screen.size_hint_x = 0.45
            self.btn_loc_screen.opacity = 1
            self.btn_loc_screen.disabled = False
            self.btn_loc_screen.size_hint_x = 0.45

        rv_data = []
        for it in self.cart:
            q_disp = FinancialEngine.format_qty(it.get('qty', 0))
            if is_qty_only:
                det_txt = f'Qté : {q_disp}'
                d_col = [0.1, 0.4, 0.8, 1]
            else:
                line_calc = FinancialEngine.calculate_document_totals([it], doc_type=doc_type)
                p = FinancialEngine.to_decimal(it.get('price', 0))
                t_rate = FinancialEngine.to_decimal(it.get('tva', 0)) if is_invoice else Decimal('0')
                det_txt = f'{FinancialEngine.format_money(p)} x {q_disp}'
                if t_rate > 0:
                    det_txt += f' (+{int(t_rate)}% TVA)'
                det_txt += f" = {FinancialEngine.format_money(line_calc['final_total'])}"
                d_col = [0.35, 0.35, 0.35, 1]
            rv_data.append({'name': it.get('name', 'Article'), 'details': det_txt, 'd_color': d_col, 'raw_item': it})

        self.rv_cart.data = rv_data
        self.rv_cart.refresh_from_data()

    def edit_cart_item(self, item: dict):
        doc_type = self.editing_doc_type or AppConstants.get_doc_type_for_mode(self.current_mode)
        is_invoice = doc_type in AppConstants.INVOICE_TAX_DOC_TYPES
        is_qty_only = doc_type in AppConstants.QTY_ONLY_DOC_TYPES
        is_virtual = AppConstants.is_default_product(item)

        dialog_h = dp(580 if is_invoice else (360 if is_qty_only else 500))
        content = MDBoxLayout(orientation='vertical', spacing='10dp', size_hint_y=None, height=dialog_h, padding=[0, '5dp', 0, 0])
        self.active_edit_target = 'qty'
        self.input_reset_mode = True

        def _upd_colors():
            act, inact = (0.9, 1, 0.9, 1), (0.95, 0.95, 0.95, 1)
            for attr, key in (('edit_qty_card', 'qty'), ('edit_price_card', 'price'), ('edit_tva_card', 'tva'), ('edit_name_card', 'name')):
                card = getattr(self, attr, None)
                if card:
                    card.md_bg_color = act if self.active_edit_target == key else inact
                    card.elevation = 3 if self.active_edit_target == key else 0

        if is_virtual:
            self.edit_name_card = MDCard(size_hint_y=None, height='68dp', radius=[10], padding=[10, 0], elevation=0)
            self.edit_name_field = SmartTextField(text=item.get('name', AppConstants.DEFAULT_PRODUCT_NAME), hint_text="Nom de l'article", font_size='20sp', halign='center', mode='line', line_color_normal=(0, 0, 0, 0), line_color_focus=(0, 0, 0, 0))
            self.edit_name_field.bind(on_touch_down=lambda i, t: (setattr(self, 'active_edit_target', 'name'), _upd_colors(), False)[2] if i.collide_point(*t.pos) else False)
            self.edit_name_card.add_widget(self.edit_name_field)
            n_row = MDBoxLayout(size_hint_y=None, height='72dp', padding=[20, 0])
            n_row.add_widget(self.edit_name_card)
            content.add_widget(n_row)
        else:
            if hasattr(self, 'edit_name_field'):
                del self.edit_name_field
            if hasattr(self, 'edit_name_card'):
                del self.edit_name_card
            content.add_widget(MDLabel(text=self.fix_text(item.get('name', 'Article')), halign='center', bold=True, font_style='Subtitle1', theme_text_color='Primary', adaptive_height=True))

        if not is_qty_only:
            self.edit_price_card = MDCard(size_hint_y=None, height='68dp', radius=[10], padding=[10, 0], elevation=0)
            self.edit_price_field = NoMenuTextField(text=FinancialEngine.format_qty(item.get('price', 0)), hint_text='Prix (DA)', font_size='26sp', halign='center', mode='line', readonly=True, line_color_normal=(0, 0, 0, 0), line_color_focus=(0, 0, 0, 0))
            self.edit_price_field.theme_text_color = 'Custom'
            self.edit_price_field.text_color_normal = (0, 0, 0, 1)
            self.edit_price_field.text_color_focus = (0, 0, 0, 1)
            self.edit_price_field.bind(on_touch_down=lambda i, t: (setattr(self, 'input_reset_mode', self.active_edit_target != 'price' or self.input_reset_mode), setattr(self, 'active_edit_target', 'price'), _upd_colors(), True)[3] if i.collide_point(*t.pos) else False)
            self.edit_price_card.add_widget(self.edit_price_field)
            pr_row = MDBoxLayout(size_hint_y=None, height='72dp', padding=[60, 0])
            pr_row.add_widget(self.edit_price_card)
            content.add_widget(pr_row)
        else:
            if hasattr(self, 'edit_price_field'):
                del self.edit_price_field
            if hasattr(self, 'edit_price_card'):
                del self.edit_price_card

        if is_invoice:
            self.edit_tva_card = MDCard(size_hint_y=None, height='58dp', radius=[10], padding=[10, 0], elevation=0)
            self.edit_tva_field = NoMenuTextField(text=FinancialEngine.format_qty(item.get('tva', 0)), hint_text='TVA %', font_size='24sp', halign='center', mode='line', readonly=True, line_color_normal=(0, 0, 0, 0), line_color_focus=(0, 0, 0, 0))
            self.edit_tva_field.theme_text_color = 'Custom'
            self.edit_tva_field.text_color_normal = (0, 0, 0, 1)
            self.edit_tva_field.text_color_focus = (0, 0, 0, 1)
            self.edit_tva_field.bind(on_touch_down=lambda i, t: (setattr(self, 'input_reset_mode', self.active_edit_target != 'tva' or self.input_reset_mode), setattr(self, 'active_edit_target', 'tva'), _upd_colors(), True)[3] if i.collide_point(*t.pos) else False)
            self.edit_tva_card.add_widget(self.edit_tva_field)
            tv_row = MDBoxLayout(size_hint_y=None, height='62dp', padding=[100, 0])
            tv_row.add_widget(self.edit_tva_card)
            content.add_widget(tv_row)
        else:
            if hasattr(self, 'edit_tva_field'):
                del self.edit_tva_field
            if hasattr(self, 'edit_tva_card'):
                del self.edit_tva_card

        qty_row = MDBoxLayout(orientation='horizontal', spacing='10dp', size_hint_y=None, height='65dp', padding=[40, 0])
        btn_minus = MDIconButton(icon='minus', theme_text_color='Custom', text_color=(1, 1, 1, 1), md_bg_color=(0.9, 0.3, 0.3, 1), pos_hint={'center_y': 0.5})
        self.edit_qty_card = MDCard(size_hint_x=1, size_hint_y=None, height='60dp', radius=[10], padding=[10, 0], elevation=0, pos_hint={'center_y': 0.5})
        self.edit_qty_field = NoMenuTextField(text=FinancialEngine.format_qty(item.get('qty', 1)), hint_text='Qté', font_size='28sp', halign='center', readonly=True, mode='line', line_color_normal=(0, 0, 0, 0), line_color_focus=(0, 0, 0, 0))
        self.edit_qty_field.theme_text_color = 'Custom'
        self.edit_qty_field.text_color_normal = (0, 0, 0, 1)
        self.edit_qty_field.text_color_focus = (0, 0, 0, 1)
        self.edit_qty_field.bind(on_touch_down=lambda i, t: (setattr(self, 'input_reset_mode', self.active_edit_target != 'qty' or self.input_reset_mode), setattr(self, 'active_edit_target', 'qty'), _upd_colors(), True)[3] if i.collide_point(*t.pos) else False)
        self.edit_qty_card.add_widget(self.edit_qty_field)
        btn_plus = MDIconButton(icon='plus', theme_text_color='Custom', text_color=(1, 1, 1, 1), md_bg_color=(0.2, 0.7, 0.2, 1), pos_hint={'center_y': 0.5})
        qty_row.add_widget(btn_minus)
        qty_row.add_widget(self.edit_qty_card)
        qty_row.add_widget(btn_plus)
        content.add_widget(qty_row)

        self.btn_save_edit = MDRaisedButton(text='MODIFIER', md_bg_color=(0, 0.6, 0.2, 1), text_color=(1, 1, 1, 1), size_hint_x=0.7, size_hint_y=1, font_size='18sp', elevation=3)

        def _upd_calc():
            if is_qty_only:
                self.btn_save_edit.text = 'MODIFIER'
                return
            q = FinancialEngine.to_decimal(self.edit_qty_field.text or '0')
            p = FinancialEngine.to_decimal(self.edit_price_field.text or '0') if hasattr(self, 'edit_price_field') else Decimal('0')
            tva = FinancialEngine.to_decimal(self.edit_tva_field.text or '0') if hasattr(self, 'edit_tva_field') else Decimal('0')
            calc = FinancialEngine.calculate_document_totals([{'id': item['id'], 'qty': q, 'price': p, 'tva': tva}], doc_type=doc_type)
            self.btn_save_edit.text = f"MODIFIER\n{FinancialEngine.format_money(calc['final_total'])}"

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

        grid = MDGridLayout(cols=3, spacing='8dp', size_hint_y=1, padding=[20, 0])
        for key in ('7', '8', '9', '4', '5', '6', '1', '2', '3', '.', '0', 'DEL'):
            if key == 'DEL':
                grid.add_widget(MDIconButton(icon='backspace-outline', theme_text_color='Custom', text_color=(0, 0, 0, 1), md_bg_color=(0.8, 0.8, 0.8, 1), size_hint=(1, 1), on_release=_bkspc))
            else:
                grid.add_widget(MDRaisedButton(text=key, md_bg_color=(0.95, 0.95, 0.95, 1), theme_text_color='Custom', text_color=(0, 0, 0, 1), font_size='22sp', size_hint=(1, 1), elevation=1, on_release=lambda x, k=key: _digit(k)))
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
            if prod_db:
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
        btns = MDBoxLayout(orientation='horizontal', spacing='10dp', size_hint_y=None, height='60dp')
        btns.add_widget(MDFlatButton(text='ANNULER', theme_text_color='Custom', text_color=(0.5, 0.5, 0.5, 1), size_hint_x=0.3, on_release=lambda x: self.edit_dialog.dismiss()))
        btns.add_widget(self.btn_save_edit)
        content.add_widget(btns)

        _upd_colors()
        _upd_calc()
        self.edit_dialog = MDDialog(title='', type='custom', content_cls=content, buttons=[], size_hint=(0.85, None))
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
            self.dialog = MDDialog(title='Attention', text='Le panier est vide !', buttons=[MDFlatButton(text='OK', on_release=lambda x: self.dialog.dismiss())])
            self.dialog.open()
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

        if fin_f == 0 or doc_type in AppConstants.RETURN_DOC_TYPES or is_comptoir:
            method_val = AppConstants.PAY_DEFAULT if (fin_f == 0 or doc_type in AppConstants.RETURN_DOC_TYPES or doc_type in AppConstants.STAMP_DUTY_DOC_TYPES) else AppConstants.PAY_CASH
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

        content = MDBoxLayout(
            orientation='vertical',
            spacing='10dp',
            size_hint_y=None,
            height=dp(640 if self.show_tax_details else 580),
            padding=['10dp', '0dp', '10dp', '20dp'],
        )
        header_box = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(50), spacing=dp(10))
        header_box.add_widget(MDLabel(text='Paiement', font_style='H5', bold=True, theme_text_color='Primary', valign='center'))
        if self.is_facture_tax:
            self.btn_payment_method = MDRaisedButton(
                text=self.payment_methods[self.current_method_index]['label'],
                md_bg_color=(0.2, 0.2, 0.2, 1),
                elevation=2,
                pos_hint={'center_y': 0.5},
                on_release=self._cycle_payment_method,
            )
            header_box.add_widget(self.btn_payment_method)
        content.add_widget(header_box)

        total_card = MDCard(orientation='vertical', size_hint_y=None, height=dp(125 if self.show_tax_details else 90), radius=[10], md_bg_color=(0.95, 0.95, 0.95, 1), elevation=1, padding='5dp', ripple_behavior=True)
        total_card.bind(on_release=lambda inst: [setattr(self.txt_paid, 'text', f'{self.current_final_total:.2f}'), self._recalc_ui_totals()])

        if self.show_tax_details:
            det_box = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(25), padding=[dp(5), 0], spacing=dp(5))
            self.lbl_pay_ht = MDLabel(text='', theme_text_color='Secondary', font_style='Caption', halign='left', bold=True, size_hint_x=0.33)
            self.lbl_timbre = MDLabel(text='', theme_text_color='Custom', text_color=(0.5, 0, 0.5, 1), font_style='Caption', halign='center', bold=True, size_hint_x=0.33)
            self.lbl_pay_tva = MDLabel(text='', theme_text_color='Custom', text_color=(0.8, 0, 0, 1), font_style='Caption', halign='right', bold=True, size_hint_x=0.33)
            det_box.add_widget(self.lbl_pay_ht)
            det_box.add_widget(self.lbl_timbre)
            det_box.add_widget(self.lbl_pay_tva)
            total_card.add_widget(det_box)
        else:
            self.lbl_pay_ht = None
            self.lbl_timbre = None
            self.lbl_pay_tva = None

        tot_box = MDBoxLayout(orientation='vertical', spacing=0)
        tot_box.add_widget(MDLabel(text='NET À PAYER', halign='center', font_style='Caption', theme_text_color='Secondary', size_hint_y=None, height=dp(20)))
        self.lbl_final_total = MDLabel(text='', halign='center', font_style='H4', bold=True, theme_text_color='Primary')
        tot_box.add_widget(self.lbl_final_total)
        total_card.add_widget(tot_box)
        content.add_widget(total_card)

        self.txt_paid = MDTextField(
            text=f'{FinancialEngine.quantize_money(self.editing_payment_amount):.2f}',
            hint_text='Montant Versé (DA)',
            font_size='40sp',
            halign='center',
            readonly=True,
            size_hint_y=None,
            height=dp(80),
            mode='fill',
            line_color_focus=(0, 0, 0, 0),
        )
        content.add_widget(self.txt_paid)

        self.lbl_rest = MDLabel(text='', halign='center', theme_text_color='Custom', font_style='H6', bold=True, size_hint_y=None, height=dp(30))
        content.add_widget(self.lbl_rest)

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
            if self.txt_paid.text:
                self.txt_paid.text = self.txt_paid.text[:-1]
            if not self.txt_paid.text:
                self.txt_paid.text = '0'
            self._recalc_ui_totals()

        grid = MDGridLayout(cols=3, spacing='10dp', size_hint_y=1)
        for key in ('7', '8', '9', '4', '5', '6', '1', '2', '3', '.', '0', 'DEL'):
            if key == 'DEL':
                grid.add_widget(MDIconButton(icon='backspace', theme_text_color='Custom', text_color=(1, 1, 1, 1), md_bg_color=(0.4, 0.4, 0.4, 1), size_hint=(1, 1), on_release=_bkspc))
            else:
                grid.add_widget(MDRaisedButton(text=key, md_bg_color=(1, 1, 1, 1), theme_text_color='Custom', text_color=(0, 0, 0, 1), font_size='24sp', size_hint=(1, 1), elevation=1, on_release=lambda x, k=key: _digit(k)))
        content.add_widget(grid)
        content.add_widget(MDBoxLayout(size_hint_y=None, height='12dp'))

        btns = MDBoxLayout(orientation='horizontal', spacing='10dp', size_hint_y=None, height='55dp')
        btns.add_widget(MDFlatButton(text='ANNULER', theme_text_color='Custom', text_color=(0.5, 0.5, 0.5, 1), size_hint_x=0.3, on_release=lambda x: self.pay_dialog.dismiss()))
        btns.add_widget(MDRaisedButton(text='VALIDER', md_bg_color=(0, 0.7, 0.2, 1), text_color=(1, 1, 1, 1), size_hint_x=0.7, size_hint_y=1, font_size='20sp', elevation=2, on_release=lambda x: self.finalize_submission(self.current_final_total)))
        content.add_widget(btns)

        self.pay_dialog = MDDialog(title='', type='custom', content_cls=content, buttons=[], size_hint=(0.94, 0.98))
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
        if diff >= 0:
            self.lbl_rest.text = f'RESTE : {FinancialEngine.format_money(diff)}'
            self.lbl_rest.text_color = (0.8, 0, 0, 1)
        else:
            self.lbl_rest.text = f'RENDU : {FinancialEngine.format_money(abs(diff))}'
            self.lbl_rest.text_color = (0, 0.6, 0.2, 1)

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
        content = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True, spacing='14dp', padding=[0, '10dp', 0, 0])
        content.add_widget(MDLabel(text=f'[b]Montant saisi :[/b] {FinancialEngine.format_money(paid)}\n[b]Total :[/b] {FinancialEngine.format_money(total)}', markup=True, halign='center', adaptive_height=True))
        content.add_widget(MDLabel(text=f"L'excédent [color=#00C853][b]({FinancialEngine.format_money(excess)})[/b][/color] sera enregistré comme une opération séparée [b](VERSEMENT / RÈGLEMENT)[/b].", markup=True, halign='center', adaptive_height=True))
        self.overpay_dialog = MDDialog(
            title="Création d'un versement",
            type='custom',
            content_cls=content,
            buttons=[
                MDFlatButton(text='CORRIGER', on_release=lambda x: [self.overpay_dialog.dismiss(), self.open_payment_dialog(None)]),
                MDRaisedButton(text='CONFIRMER', md_bg_color=(0, 0.7, 0.2, 1), on_release=lambda x: [self.overpay_dialog.dismiss(), self.process_transaction(paid, total, method=method)]),
            ],
        )
        self.overpay_dialog.open()

    def show_credit_warning(self, paid: Decimal, total: Decimal, remaining: Decimal, method: str):
        content = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True, spacing='14dp', padding=[0, '10dp', 0, 0])
        content.add_widget(MDLabel(text=f'[b]Montant saisi :[/b] {FinancialEngine.format_money(paid)}\n[b]Total :[/b] {FinancialEngine.format_money(total)}', markup=True, halign='center', adaptive_height=True))
        content.add_widget(MDLabel(text=f'Le montant versé est insuffisant.\nLe reste [color=#D32F2F][b]({FinancialEngine.format_money(remaining)})[/b][/color] sera enregistré en [b]CRÉDIT[/b].', markup=True, halign='center', adaptive_height=True))
        self.debt_dialog = MDDialog(
            title='Attention : Crédit',
            type='custom',
            content_cls=content,
            buttons=[
                MDFlatButton(text='ANNULER', on_release=lambda x: self.debt_dialog.dismiss()),
                MDRaisedButton(text='CONFIRMER', md_bg_color=(0.8, 0, 0, 1), on_release=lambda x: [self.debt_dialog.dismiss(), self.process_transaction(paid, total, method=method)]),
            ],
        )
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
            if fin_factor == 1 and inv_paid > real_total:
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
                'location': self.selected_location,
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
            if self.db.get_bool_setting('printer_auto', False) and self.db.get_setting('printer_mac', ''):
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
    # 7.8 CLIENTS & FOURNISSEURS (SÉLECTION, GESTION, GPS, VERSEMENTS)
    # --------------------------------------------------------------------------
    def show_entity_selection_dialog(self, x, next_action=None):
        self.pending_entity_next_action = next_action
        e_type = AppConstants.get_entity_type(self.current_mode)
        self.entities_source_type = e_type
        title = 'Choisir un Fournisseur' if e_type == AppConstants.ENTITY_SUPPLIER else 'Choisir un Client'

        content = MDBoxLayout(orientation='vertical', size_hint_y=None, height=dp(580))
        self.entity_search = SmartTextField(hint_text='Rechercher...', icon_right='magnify')
        self.entity_search.bind(text=self.filter_entities_paginated)
        content.add_widget(self.entity_search)
        self.rv_entity = EntityRecycleView()
        content.add_widget(self.rv_entity)

        self.entity_dialog = MDDialog(title=title, type='custom', content_cls=content, size_hint=(0.9, 0.8))
        self.entity_dialog.open()
        self.active_entity_rv = self.rv_entity
        self.load_more_entities(reset=True)

    def open_entity_manager(self, entity_type: str):
        self.current_entity_type_mgmt = entity_type
        self.current_entity_sort = 'name'
        title = 'Gestion Clients' if entity_type == AppConstants.ENTITY_CLIENT else 'Gestion Fournisseurs'

        content = MDBoxLayout(orientation='vertical', size_hint_y=None, height=dp(600), spacing=dp(5))
        search_row = MDBoxLayout(orientation='horizontal', spacing=dp(5), size_hint_y=None, height=dp(55))
        self.entity_search = SmartTextField(hint_text='Rechercher...', icon_right='magnify', size_hint_x=0.85)
        self.entity_search.bind(text=self.filter_entities_paginated)
        self.btn_sort_entity = MDIconButton(icon='sort-alphabetical-variant', theme_text_color='Custom', text_color=(0.2, 0.2, 0.2, 1), md_bg_color=(0.95, 0.95, 0.95, 1), size_hint=(None, None), size=(dp(48), dp(48)), pos_hint={'center_y': 0.5}, on_release=self.toggle_entity_sort)
        search_row.add_widget(self.entity_search)
        search_row.add_widget(self.btn_sort_entity)
        content.add_widget(search_row)

        self.rv_mgmt_entity = MgmtEntityRecycleView()
        content.add_widget(self.rv_mgmt_entity)
        content.add_widget(MDFillRoundFlatButton(text='AJOUTER NOUVEAU', size_hint_x=1, md_bg_color=(0, 0.7, 0.2, 1), on_release=lambda x: self.show_add_edit_entity_dialog(None)))

        self.mgmt_dialog = MDDialog(title=title, type='custom', content_cls=content, size_hint=(0.95, 0.9))
        self.mgmt_dialog.open()
        self.active_entity_rv = self.rv_mgmt_entity
        self.load_more_entities(reset=True)

    def toggle_entity_sort(self, instance):
        if self.current_entity_sort == 'name':
            self.current_entity_sort = 'balance'
            instance.icon = 'sort-numeric-descending'
            instance.md_bg_color = (0.8, 0.9, 1, 1)
            self.notify('Tri : Par solde décroissant', 'info')
        elif self.current_entity_sort == 'balance':
            self.current_entity_sort = 'active'
            instance.icon = 'check-bold'
            instance.md_bg_color = (0.8, 1, 0.8, 1)
            self.notify("Tri : Opérations d'aujourd'hui d'abord", 'info')
        else:
            self.current_entity_sort = 'name'
            instance.icon = 'sort-alphabetical-variant'
            instance.md_bg_color = (0.95, 0.95, 0.95, 1)
            self.notify('Tri : Alphabétique (A-Z)', 'info')
        self.load_more_entities(reset=True)

    def filter_entities_paginated(self, instance, text=None):
        if self._entity_search_event:
            self._entity_search_event.cancel()
        self._entity_search_event = Clock.schedule_once(lambda dt: self.load_more_entities(reset=True), 0.3)

    def load_more_entities(self, reset: bool = False):
        if self.is_loading_entities and not reset:
            return
        is_mgmt = bool(getattr(self, 'mgmt_dialog', None) and self.mgmt_dialog.parent)
        is_sel = bool(getattr(self, 'entity_dialog', None) and self.entity_dialog.parent)
        if not is_mgmt and not is_sel:
            return

        target_type = self.current_entity_type_mgmt if is_mgmt else getattr(self, 'entities_source_type', AppConstants.ENTITY_CLIENT)
        self.active_entity_rv = self.rv_mgmt_entity if is_mgmt else self.rv_entity
        search_q = self.entity_search.get_value().strip() if hasattr(self, 'entity_search') else ''

        if reset:
            self.entity_page_offset = 0
            self.is_loading_entities = False
            if self.active_entity_rv:
                self.active_entity_rv.scroll_y = 1.0
                self.active_entity_rv.data = []

        self.is_loading_entities = True
        active_today = self.db.get_active_entity_ids_today(target_type)
        rows = self.db.get_entities(target_type, search_query=search_q, limit=50, offset=self.entity_page_offset, sort_by=self.current_entity_sort if is_mgmt else 'name', active_ids=active_today)

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
                rv_data.append({'raw_name': e['name'], 'balance_text': '[size=14sp][b][color=404040]-- Direct --[/color][/b][/size]', 'icon': 'store', 'icon_color': [0, 0.4, 0.7, 1], 'raw_data': e})
            else:
                col_hex = '00C853' if is_client else 'D50000'
                rv_data.append({
                    'raw_name': e['name'],
                    'balance_text': f'Solde : [color={col_hex}][b]{bal_str}[/b][/color]',
                    'icon': 'account-check' if bal <= 0 else 'account-alert',
                    'icon_color': [0, 0.7, 0.2, 1] if bal <= 0 else [0.9, 0, 0, 1],
                    'raw_data': e,
                })

        if self.active_entity_rv:
            if reset:
                self.active_entity_rv.data = rv_data
                self.entity_page_offset = len(rows)
            else:
                self.active_entity_rv.data.extend(rv_data)
                self.entity_page_offset += len(rows)
            self.active_entity_rv.refresh_from_data()
            self.active_entity_rv.loading_lock = False
        self.is_loading_entities = False

    def select_entity_from_rv(self, entity_data: dict):
        self.selected_entity = {
            'id': entity_data['id'],
            'name': entity_data['name'],
            'category': AppConstants.normalize_price_category(entity_data.get('price_category', AppConstants.PRICE_CAT_DETAIL)),
        }
        if hasattr(self, 'btn_ent_screen'):
            self.btn_ent_screen.text = self.fix_text(entity_data['name'])[:15]
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
        gps_data = entity.get('gps_location', '')
        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, height=dp(160), padding=[dp(15), 0, dp(15), dp(10)])
        content.add_widget(
            MDFillRoundFlatIconButton(
                text='Localiser (Maps)',
                icon='google-maps',
                size_hint_x=1,
                height=dp(50),
                md_bg_color=(0.1, 0.7, 0.3, 1) if gps_data else (0.7, 0.7, 0.7, 1),
                theme_text_color='Custom',
                text_color=(1, 1, 1, 1),
                icon_color=(1, 1, 1, 1),
                on_release=lambda x: [self.options_dialog.dismiss(), self.open_client_location(gps_data)],
            )
        )
        content.add_widget(MDSeparator(height=dp(1), color=(0.9, 0.9, 0.9, 1)))
        row = MDBoxLayout(orientation='horizontal', spacing=dp(10), size_hint_y=None, height=dp(50))
        row.add_widget(MDFillRoundFlatIconButton(text='Modifier', icon='pencil', size_hint_x=0.5, height=dp(48), md_bg_color=(0.15, 0.45, 0.8, 1), theme_text_color='Custom', text_color=(1, 1, 1, 1), icon_color=(1, 1, 1, 1), on_release=lambda x: [self.options_dialog.dismiss(), self.show_add_edit_entity_dialog(entity)]))
        row.add_widget(MDRoundFlatIconButton(text='Supprimer', icon='delete', size_hint_x=0.5, height=dp(48), theme_text_color='Custom', text_color=(0.85, 0.1, 0.1, 1), icon_color=(0.85, 0.1, 0.1, 1), line_color=(0.85, 0.1, 0.1, 1), on_release=lambda x: self.confirm_delete_entity(entity)))
        content.add_widget(row)
        self.options_dialog = MDDialog(title=self.fix_text(entity['name']), type='custom', content_cls=content, size_hint=(0.95, None), radius=[16, 16, 16, 16])
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
        scroll = MDScrollView(size_hint_y=None, height=dp(580))
        main_box = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(14), padding=[dp(10), dp(10), dp(10), dp(20)])

        card_id = MDCard(orientation='vertical', radius=[12], padding=dp(15), spacing=dp(10), elevation=1, adaptive_height=True)
        card_id.add_widget(MDLabel(text='Identité', bold=True, theme_text_color='Primary', font_style='Subtitle1'))
        f_name = SmartTextField(text=entity.get('name', '') if is_edit else '', hint_text='Nom Complet / Raison Sociale *', required=True, icon_right='account')
        f_act = SmartTextField(text=entity.get('activity', '') if is_edit else '', hint_text='Activité', icon_right='briefcase')
        card_id.add_widget(f_name)
        card_id.add_widget(f_act)
        main_box.add_widget(card_id)

        card_ct = MDCard(orientation='vertical', radius=[12], padding=dp(15), spacing=dp(10), elevation=1, adaptive_height=True, md_bg_color=(0.96, 0.98, 1, 1))
        card_ct.add_widget(MDLabel(text='Coordonnées & GPS', bold=True, theme_text_color='Primary', font_style='Subtitle1'))
        f_phone = SmartTextField(text=entity.get('phone', '') if is_edit else '', hint_text='Téléphone', icon_right='phone')
        f_addr = SmartTextField(text=entity.get('address', '') if is_edit else '', hint_text='Adresse', icon_right='map-marker')
        f_email = SmartTextField(text=entity.get('email', '') if is_edit else '', hint_text='Email', icon_right='email')
        gps_row = MDBoxLayout(orientation='horizontal', spacing=dp(5), adaptive_height=True)
        f_gps = SmartTextField(text=entity.get('gps_location', '') if is_edit else '', hint_text='Lien Maps / Position GPS', icon_right='google-maps', size_hint_x=0.85)
        gps_row.add_widget(f_gps)
        gps_row.add_widget(MDIconButton(icon='content-paste', theme_text_color='Custom', text_color=self.theme_cls.primary_color, on_release=lambda x: f_gps.set_raw_value(Clipboard.paste())))
        for w in (f_phone, f_addr, gps_row, f_email):
            card_ct.add_widget(w)
        main_box.add_widget(card_ct)

        disp_cat = AppConstants.normalize_price_category(entity.get('price_category', AppConstants.PRICE_CAT_DETAIL) if is_edit else AppConstants.PRICE_CAT_DETAIL)
        f_cat = MDTextField(text=disp_cat, hint_text='Catégorie de Prix', readonly=True, icon_right='tag')
        f_cat.bind(on_touch_down=lambda i, t: (self.show_price_cat_selector(i), True)[1] if i.collide_point(*t.pos) else False)
        if self.current_entity_type_mgmt == AppConstants.ENTITY_CLIENT:
            card_cm = MDCard(orientation='vertical', radius=[12], padding=dp(15), spacing=dp(10), elevation=1, adaptive_height=True)
            card_cm.add_widget(MDLabel(text='Commercial', bold=True, theme_text_color='Primary', font_style='Subtitle1'))
            card_cm.add_widget(f_cat)
            main_box.add_widget(card_cm)

        card_fisc = MDCard(orientation='vertical', radius=[12], padding=dp(15), spacing=dp(10), elevation=1, adaptive_height=True, md_bg_color=(0.95, 0.95, 0.95, 1))
        card_fisc.add_widget(MDLabel(text='Informations Fiscales', bold=True, theme_text_color='Primary', font_style='Subtitle1'))
        f_rc = SmartTextField(text=entity.get('rc', '') if is_edit else '', hint_text='N° Registre Commerce (RC)')
        f_nif = SmartTextField(text=entity.get('nif', '') if is_edit else '', hint_text='N.I.F')
        f_nis = SmartTextField(text=entity.get('nis', '') if is_edit else '', hint_text='N.I.S')
        f_nai = SmartTextField(text=entity.get('nai', '') if is_edit else '', hint_text='N.A.I')
        for w in (f_rc, f_nif, f_nis, f_nai):
            card_fisc.add_widget(w)
        main_box.add_widget(card_fisc)

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
                })
                self.ae_dialog.dismiss()
                self.notify('Enregistré avec succès.', 'success')
                self.load_more_entities(reset=True)
            except Exception as exc:
                self.notify(str(exc), 'error')

        footer = MDBoxLayout(orientation='vertical', spacing=dp(10), adaptive_height=True, padding=[0, dp(10), 0, 0])
        footer.add_widget(MDRaisedButton(text='ENREGISTRER', md_bg_color=(0, 0.7, 0.2, 1), text_color=(1, 1, 1, 1), size_hint_x=1, height=dp(50), on_release=_save))
        footer.add_widget(MDRaisedButton(text='ANNULER', md_bg_color=(0.9, 0.9, 0.9, 1), text_color=(0.2, 0.2, 0.2, 1), size_hint_x=1, height=dp(50), on_release=lambda x: self.ae_dialog.dismiss()))
        main_box.add_widget(footer)
        scroll.add_widget(main_box)

        self.ae_dialog = MDDialog(title='Modifier la fiche' if is_edit else 'Nouveau Tiers', type='custom', content_cls=scroll, buttons=[], size_hint=(0.96, 0.94))
        self.ae_dialog.open()
        Clock.schedule_once(lambda dt: setattr(f_name, 'focus', True), 0.3)

    def show_price_cat_selector(self, tf):
        content = MDBoxLayout(orientation='vertical', spacing=10, size_hint_y=None, height=dp(160), padding=dp(10))
        for cat, col in ((AppConstants.PRICE_CAT_DETAIL, (0, 0.6, 0.6, 1)), (AppConstants.PRICE_CAT_SEMI, (0.9, 0.6, 0, 1)), (AppConstants.PRICE_CAT_GROS, (0.5, 0, 0.5, 1))):
            content.add_widget(MDRaisedButton(text=cat, size_hint_x=1, md_bg_color=col, on_release=lambda x, v=cat: [setattr(tf, 'text', v), self.cat_dialog.dismiss()]))
        self.cat_dialog = MDDialog(title='Catégorie de Prix', type='custom', content_cls=content, size_hint=(0.8, None))
        self.cat_dialog.open()

    def confirm_delete_entity(self, entity: dict):
        if getattr(self, 'options_dialog', None):
            self.options_dialog.dismiss()

        def _do_del(x):
            if getattr(self, 'del_conf_dialog', None):
                self.del_conf_dialog.dismiss()
            try:
                if self.db.delete_entity(entity['id']):
                    self.notify('Supprimé avec succès.', 'success')
                    self.load_more_entities(reset=True)
                else:
                    self.play_sound('error')
                    self.notify("Suppression impossible : ce compte possède des opérations.", 'error')
            except Exception as exc:
                self.notify(str(exc), 'error')

        self.del_conf_dialog = MDDialog(
            title='Confirmation',
            text=f"Voulez-vous vraiment supprimer {self.fix_text(entity['name'])} ?",
            buttons=[
                MDFlatButton(text='NON', on_release=lambda x: self.del_conf_dialog.dismiss()),
                MDRaisedButton(text='OUI, SUPPRIMER', md_bg_color=(0.8, 0, 0, 1), on_release=_do_del),
            ],
        )
        self.del_conf_dialog.open()

    def show_simple_payment_dialog(self, amount=None):
        if not self.selected_entity:
            return
        self.temp_note = getattr(self, 'temp_note', '')
        is_client = self.current_mode in ('client_payment', 'client_pay')
        title = 'Versement Client' if is_client else 'Règlement Fournisseur'
        theme_col = (0, 0.6, 0.2, 1) if is_client else (0.8, 0.4, 0, 1)

        content = MDBoxLayout(orientation='vertical', spacing='12dp', size_hint_y=None, height=dp(560), padding=[0, '10dp', 0, 0])
        content.add_widget(MDLabel(text=self.fix_text(self.selected_entity['name']), halign='center', font_style='H5', bold=True, theme_text_color='Primary', adaptive_height=True))

        def _bkspc(x=None):
            if self.txt_simple_amount.text:
                self.txt_simple_amount.text = self.txt_simple_amount.text[:-1]

        def _digit(d: str):
            cur = self.txt_simple_amount.text
            if d == '.':
                if '.' not in cur:
                    self.txt_simple_amount.text = '0.' if not cur else cur + '.'
            elif d == '-':
                self.txt_simple_amount.text = cur[1:] if cur.startswith('-') else '-' + cur
            elif cur == '0':
                self.txt_simple_amount.text = d
            elif cur == '-0':
                self.txt_simple_amount.text = '-' + d
            else:
                self.txt_simple_amount.text = cur + d

        val_str = FinancialEngine.format_qty(amount) if amount is not None else ''
        in_row = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(110), spacing='10dp', padding=[dp(20), 0])
        self.txt_simple_amount = MDTextField(text=val_str, hint_text='Montant (DA)', font_size='42sp', halign='center', readonly=True, mode='fill', line_color_focus=theme_col, size_hint_x=0.7)
        self.btn_note_icon = MDIconButton(icon='note-edit-outline', theme_text_color='Custom', text_color=(1, 1, 1, 1), md_bg_color=(0, 0.6, 0.2, 1) if self.temp_note else (0.2, 0.2, 0.2, 1), size_hint=(None, None), size=(dp(55), dp(55)), pos_hint={'center_y': 0.5}, on_release=self.open_note_input)
        in_row.add_widget(self.txt_simple_amount)
        in_row.add_widget(self.btn_note_icon)
        in_row.add_widget(MDIconButton(icon='backspace-outline', theme_text_color='Custom', text_color=(1, 1, 1, 1), md_bg_color=(0.9, 0.1, 0.1, 1), size_hint=(None, None), size=(dp(55), dp(55)), pos_hint={'center_y': 0.5}, on_release=_bkspc))
        content.add_widget(in_row)

        grid = MDGridLayout(cols=3, spacing='10dp', size_hint_y=1, padding=[dp(20), dp(10)])
        for key in ('7', '8', '9', '4', '5', '6', '1', '2', '3', '-', '0', '.'):
            if key == '-':
                btn = MDCard(md_bg_color=(1, 0.85, 0.85, 1), elevation=1, size_hint=(1, 1), radius=[4], ripple_behavior=True, on_release=lambda x, k=key: _digit(k))
                fc = MDFloatLayout()
                cb = MDBoxLayout(orientation='vertical', adaptive_size=True, pos_hint={'center_x': 0.5, 'center_y': 0.5})
                cb.add_widget(MDIcon(icon='minus-thick', font_size='30sp', halign='center', theme_text_color='Custom', text_color=(0.8, 0, 0, 1)))
                cb.add_widget(MDLabel(text='Crédit Initial', font_style='Caption', halign='center', theme_text_color='Custom', text_color=(0.8, 0, 0, 1), adaptive_size=True))
                fc.add_widget(cb)
                btn.add_widget(fc)
                grid.add_widget(btn)
            else:
                grid.add_widget(MDRaisedButton(text=key, md_bg_color=(0.96, 0.96, 0.96, 1), theme_text_color='Custom', text_color=(0.1, 0.1, 0.1, 1), font_size='28sp', elevation=1, size_hint=(1, 1), on_release=lambda x, k=key: _digit(k)))
        content.add_widget(grid)

        btns = MDBoxLayout(orientation='horizontal', spacing='10dp', size_hint_y=None, height='70dp', padding=[0, '10dp', 0, 0])
        btns.add_widget(MDFlatButton(text='ANNULER', theme_text_color='Custom', text_color=(0.5, 0.5, 0.5, 1), size_hint_x=0.25, on_release=lambda x: self.simple_pay_dialog.dismiss()))
        btns.add_widget(MDRaisedButton(text='VALIDER', md_bg_color=theme_col, text_color=(1, 1, 1, 1), size_hint_x=0.75, size_hint_y=1, font_size='22sp', elevation=3, on_release=self.submit_simple_payment))
        content.add_widget(btns)

        self.simple_pay_dialog = MDDialog(title=title, type='custom', content_cls=content, size_hint=(0.92, None), buttons=[])
        self.simple_pay_dialog.open()

    def open_note_input(self, instance):
        content = MDBoxLayout(orientation='vertical', spacing='10dp', size_hint_y=None, height=dp(100), padding=dp(10))
        nf = SmartTextField(text=self.temp_note, hint_text='Note (optionnel)')
        content.add_widget(nf)

        def _save_n(x):
            self.temp_note = nf.get_value().strip()
            self.btn_note_icon.md_bg_color = (0, 0.6, 0.2, 1) if self.temp_note else (0.2, 0.2, 0.2, 1)
            ndlg.dismiss()

        ndlg = MDDialog(title='Ajouter une note', type='custom', content_cls=content, buttons=[MDFlatButton(text='ANNULER', on_release=lambda x: ndlg.dismiss()), MDRaisedButton(text='OK', on_release=_save_n)])
        ndlg.open()

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
            if self.db.get_bool_setting('printer_auto', False) and self.db.get_setting('printer_mac', ''):
                det = self.db.get_operation_details(saved_id, is_transfer=False)
                if det:
                    threading.Thread(target=self.print_ticket_bluetooth, args=(det,), daemon=True).start()
            self.temp_note = ''
            self.editing_tx_id = None
            self.current_editing_date = None
            self.load_more_entities(reset=True)
            self.check_and_load_stats()
            if getattr(self, 'pending_dialog', None):
                self.filter_history_list('auj')
            if getattr(self, 'entity_hist_dialog', None):
                self.filter_entity_history_list('auj')
        except Exception as exc:
            self.notify(str(exc), 'error')
        finally:
            self.is_transaction_in_progress = False

    def open_entity_history_dialog(self, entity: dict):
        self.history_target_entity = entity
        content = MDBoxLayout(orientation='vertical', size_hint_y=None, height=dp(550), spacing=dp(8))
        tabs = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(40), spacing=dp(4))
        self.btn_eh_today = MDRaisedButton(text='AUJ.', size_hint_x=0.33, elevation=0, on_release=lambda x: self.filter_entity_history_list('auj'))
        self.btn_eh_yest = MDRaisedButton(text='HIER', size_hint_x=0.33, elevation=0, md_bg_color=(0.5, 0.5, 0.5, 1), on_release=lambda x: self.filter_entity_history_list('hier'))
        self.btn_eh_cal = MDRaisedButton(text='CALENDRIER', size_hint_x=0.34, elevation=0, md_bg_color=(0.5, 0.5, 0.5, 1), on_release=self._open_entity_hist_date)
        for b in (self.btn_eh_today, self.btn_eh_yest, self.btn_eh_cal):
            tabs.add_widget(b)
        content.add_widget(tabs)
        self.rv_entity_history = HistoryRecycleView()
        content.add_widget(self.rv_entity_history)
        self.entity_hist_dialog = MDDialog(title=self.fix_text(f"Historique : {entity['name']}"), type='custom', content_cls=content, size_hint=(0.95, 0.9))
        self.entity_hist_dialog.open()
        self.filter_entity_history_list('auj')

    def _open_entity_hist_date(self, inst):
        dp_dlg = MDDatePicker(mode='range')
        dp_dlg.bind(on_save=lambda i, val, dr: self.filter_entity_history_list('custom', dr[0], dr[-1]) if dr else self.filter_entity_history_list('custom', val, val))
        dp_dlg.open()

    def filter_entity_history_list(self, ftype: str = 'auj', start_d=None, end_d=None):
        if not hasattr(self, 'rv_entity_history') or not getattr(self, 'history_target_entity', None):
            return
        inact, act = (0.5, 0.5, 0.5, 1), self.theme_cls.primary_color
        for b in (self.btn_eh_today, self.btn_eh_yest, self.btn_eh_cal):
            b.md_bg_color = inact
        today = datetime.now().date()
        if ftype == 'custom' and start_d:
            s_d, e_d = start_d, (end_d or start_d)
            self.btn_eh_cal.md_bg_color = act
            self.btn_eh_cal.text = s_d.strftime('%d/%m') if s_d == e_d else 'PÉRIODE'
        elif ftype == 'hier':
            s_d = e_d = today - timedelta(days=1)
            self.btn_eh_yest.md_bg_color = act
            self.btn_eh_cal.text = 'CALENDRIER'
        else:
            s_d = e_d = today
            self.btn_eh_today.md_bg_color = act
            self.btn_eh_cal.text = 'CALENDRIER'

        rows = self.db.get_operations_history(
            str(s_d),
            str(e_d),
            entity_id=self.history_target_entity['id'],
            entity_category=self.history_target_entity.get('entity_type'),
            limit=100,
            offset=0,
        )
        self.rv_entity_history.data = self._format_history_rows(rows)
        self.rv_entity_history.refresh_from_data()

    # --------------------------------------------------------------------------
    # 7.9 DÉPENSES / CHARGES
    # --------------------------------------------------------------------------
    def open_add_expense_dialog(self, edit_data: dict | None = None):
        if getattr(self, 'expense_dialog', None):
            try:
                self.expense_dialog.dismiss()
            except Exception:
                pass
        self.editing_expense_id = edit_data['id'] if edit_data else None
        content = MDBoxLayout(orientation='vertical', spacing=dp(16), size_hint_y=None, height=dp(290), padding=[dp(10), dp(16), dp(10), 0])
        self.exp_amount_field = MDTextField(hint_text='Montant (DA)', input_filter='float', font_size='24sp', icon_left='cash', text=str(edit_data['amount']) if edit_data else '')
        self.exp_cat_field = MDTextField(hint_text='Catégorie', text=edit_data['category'] if edit_data else 'Autre', readonly=True, icon_right='menu-down', icon_left='tag')
        self.exp_desc_field = SmartTextField(hint_text='Description / Note', text=edit_data['description'] if edit_data else '')

        def _open_cat_menu(inst, touch):
            if inst.collide_point(*touch.pos):
                scroll_cat = MDScrollView(size_hint_y=None, height=dp(320))
                cat_list = MDList()
                for c in AppConstants.EXPENSE_CATEGORIES:
                    cat_list.add_widget(
                        OneLineListItem(
                            text=c,
                            on_release=lambda x, val=c: [setattr(self.exp_cat_field, 'text', val), c_dlg.dismiss()],
                        )
                    )
                scroll_cat.add_widget(cat_list)
                c_dlg = MDDialog(
                    title='Catégorie de dépense',
                    type='custom',
                    content_cls=scroll_cat,
                    size_hint=(0.85, None),
                )
                c_dlg.open()
                return True
            return False

        self.exp_cat_field.bind(on_touch_down=_open_cat_menu)
        content.add_widget(self.exp_amount_field)
        content.add_widget(self.exp_cat_field)
        content.add_widget(self.exp_desc_field)

        if not edit_data:
            content.add_widget(
                MDFillRoundFlatIconButton(
                    text="VOIR L'HISTORIQUE",
                    icon='history',
                    md_bg_color=(0, 0, 0, 0),
                    theme_text_color='Custom',
                    text_color=self.theme_cls.primary_color,
                    icon_color=self.theme_cls.primary_color,
                    pos_hint={'center_x': 0.5},
                    on_release=lambda x: [self.expense_dialog.dismiss(), self.open_expense_history_dialog()],
                )
            )

        self.expense_dialog = MDDialog(
            title='Modifier la dépense' if edit_data else 'Nouvelle Dépense',
            type='custom',
            content_cls=content,
            size_hint=(0.9, None),
            radius=[16, 16, 16, 16],
            buttons=[
                MDFlatButton(text='ANNULER', theme_text_color='Error', on_release=lambda x: self.expense_dialog.dismiss()),
                MDRaisedButton(text='ENREGISTRER', md_bg_color=(0, 0.65, 0.25, 1), on_release=self.submit_expense),
            ],
        )
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
        if not hasattr(self, 'exp_start_date'):
            self.exp_start_date = today_str
            self.exp_end_date = today_str

        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, height=dp(520))
        frow = MDBoxLayout(orientation='horizontal', spacing=dp(5), size_hint_y=None, height=dp(45))
        self.btn_exp_start = MDFlatButton(text=self.exp_start_date, md_bg_color=(0.9, 0.9, 0.9, 1), size_hint_x=0.42, on_release=lambda x: self._pick_exp_date('start'))
        self.btn_exp_end = MDFlatButton(text=self.exp_end_date, md_bg_color=(0.9, 0.9, 0.9, 1), size_hint_x=0.42, on_release=lambda x: self._pick_exp_date('end'))
        frow.add_widget(self.btn_exp_start)
        frow.add_widget(MDLabel(text='Au', halign='center', size_hint_x=0.16, bold=True))
        frow.add_widget(self.btn_exp_end)
        content.add_widget(frow)

        scroll = MDScrollView(size_hint_y=1)
        self.exp_list_layout = MDBoxLayout(orientation='vertical', spacing=dp(8), adaptive_height=True, padding=[dp(2), dp(8)])
        scroll.add_widget(self.exp_list_layout)
        content.add_widget(scroll)

        self.lbl_exp_total = MDLabel(text='TOTAL : 0,00 DA', bold=True, font_style='H6', theme_text_color='Error', halign='right', size_hint_y=None, height=dp(30))
        content.add_widget(self.lbl_exp_total)

        self.history_exp_dialog = MDDialog(
            title='Historique des dépenses',
            type='custom',
            content_cls=content,
            size_hint=(0.95, None),
            radius=[16, 16, 16, 16],
            buttons=[
                MDFlatButton(text='+ NOUVELLE', theme_text_color='Primary', on_release=lambda x: [self.history_exp_dialog.dismiss(), self.open_add_expense_dialog()]),
                MDRaisedButton(text='FERMER', md_bg_color=(0.2, 0.2, 0.2, 1), on_release=lambda x: self.history_exp_dialog.dismiss()),
            ],
        )
        self.history_exp_dialog.open()
        self._load_expenses_ui()

    def _pick_exp_date(self, target: str):
        dp_dlg = MDDatePicker()
        dp_dlg.bind(on_save=lambda inst, val, dr: self._on_exp_date_saved(val, target))
        dp_dlg.open()

    def _on_exp_date_saved(self, val, target: str):
        if target == 'start':
            self.exp_start_date = str(val)
            self.btn_exp_start.text = str(val)
        else:
            self.exp_end_date = str(val)
            self.btn_exp_end.text = str(val)
        self._load_expenses_ui()

    def _load_expenses_ui(self):
        self.exp_list_layout.clear_widgets()
        rows = self.db.get_expenses(self.exp_start_date, self.exp_end_date)
        total = Decimal('0.00')
        for item in rows:
            amt = FinancialEngine.to_decimal(item.get('amount', 0))
            total += amt
            card = MDCard(orientation='horizontal', padding=dp(10), spacing=dp(8), size_hint_y=None, height=dp(82), radius=[12], md_bg_color=(1, 1, 1, 1), elevation=1)
            card.add_widget(MDIcon(icon='cash-minus', theme_text_color='Custom', text_color=(0.8, 0.15, 0.15, 1), font_size='28sp', size_hint_x=None, width=dp(36), pos_hint={'center_y': 0.5}))
            info = MDBoxLayout(orientation='vertical', size_hint_x=1)
            info.add_widget(MDLabel(text=f"{item['category']} - {FinancialEngine.format_money(amt)}", font_style='Subtitle2', bold=True, theme_text_color='Primary'))
            info.add_widget(MDLabel(text=self.fix_text(item['description']), font_style='Caption', theme_text_color='Secondary', shorten=True))
            info.add_widget(MDLabel(text=str(item['expense_date'])[:16], font_style='Caption', theme_text_color='Hint'))
            card.add_widget(info)
            acts = MDBoxLayout(orientation='vertical', size_hint_x=None, width=dp(40))
            acts.add_widget(MDIconButton(icon='pencil-outline', theme_text_color='Custom', text_color=(0.1, 0.5, 0.8, 1), icon_size='20sp', on_release=lambda x, i=item: [self.history_exp_dialog.dismiss(), self.open_add_expense_dialog(edit_data=i)]))
            acts.add_widget(MDIconButton(icon='trash-can-outline', theme_text_color='Custom', text_color=(0.8, 0, 0, 1), icon_size='20sp', on_release=lambda x, eid=item['id']: self.confirm_delete_expense(eid)))
            card.add_widget(acts)
            self.exp_list_layout.add_widget(card)

        self.lbl_exp_total.text = f'TOTAL : {FinancialEngine.format_money(total)}'
        if not rows:
            self.exp_list_layout.add_widget(MDLabel(text='Aucune dépense sur cette période.', halign='center', theme_text_color='Hint'))

    def confirm_delete_expense(self, exp_id: int):
        def _do_del(x):
            self.del_exp_dialog.dismiss()
            self.db.delete_expense(exp_id)
            self.notify('Dépense supprimée.', 'success')
            self._load_expenses_ui()
            self.check_and_load_stats()

        self.del_exp_dialog = MDDialog(
            title='Confirmation',
            text='Voulez-vous vraiment supprimer cette dépense ?',
            buttons=[
                MDFlatButton(text='NON', on_release=lambda x: self.del_exp_dialog.dismiss()),
                MDRaisedButton(text='SUPPRIMER', md_bg_color=(0.8, 0, 0, 1), on_release=_do_del),
            ],
        )
        self.del_exp_dialog.open()

    # --------------------------------------------------------------------------
    # 7.10 HISTORIQUE GLOBAL, DÉTAILS, MODIFICATION, SUPPRESSION & PDF A4
    # --------------------------------------------------------------------------
    def show_pending_dialog(self):
        self.current_history_filter = AppConstants.FAM_ALL
        content = MDBoxLayout(orientation='vertical', size_hint_y=None, height=dp(580), spacing=dp(8))

        tabs = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(38), spacing=dp(4))
        self.btn_hist_today = MDRaisedButton(text='AUJ', size_hint_x=1, font_size='11sp', elevation=0, on_release=lambda x: self.filter_history_list('auj'))
        self.btn_hist_yest = MDRaisedButton(text='HIE', size_hint_x=1, font_size='11sp', elevation=0, md_bg_color=(0.5, 0.5, 0.5, 1), on_release=lambda x: self.filter_history_list('hier'))
        self.btn_hist_week = MDRaisedButton(text='SEM', size_hint_x=1, font_size='11sp', elevation=0, md_bg_color=(0.5, 0.5, 0.5, 1), on_release=lambda x: self.filter_history_list('semaine'))
        self.btn_hist_month = MDRaisedButton(text='MOI', size_hint_x=1, font_size='11sp', elevation=0, md_bg_color=(0.5, 0.5, 0.5, 1), on_release=lambda x: self.filter_history_list('mois'))
        self.btn_hist_date = MDRaisedButton(text='CAL', size_hint_x=1, font_size='11sp', elevation=0, md_bg_color=(0.5, 0.5, 0.5, 1), on_release=self.open_history_date_picker)
        for b in (self.btn_hist_today, self.btn_hist_yest, self.btn_hist_week, self.btn_hist_month, self.btn_hist_date):
            tabs.add_widget(b)
        content.add_widget(tabs)

        if not self.is_seller_mode:
            self.btn_hist_filter = MDRaisedButton(text="FILTRER PAR TYPE D'OPÉRATION", size_hint_x=1, size_hint_y=None, height=dp(40), elevation=0, md_bg_color=(0.2, 0.2, 0.2, 1), on_release=self.open_history_filter_menu)
            content.add_widget(self.btn_hist_filter)

        self.rv_history = HistoryRecycleView()
        content.add_widget(self.rv_history)
        self.pending_dialog = MDDialog(title='Historique des Opérations', type='custom', content_cls=content, size_hint=(0.98, 0.96))
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
        scroll = MDScrollView(size_hint_y=None, height=dp(400))
        lst = MDList()
        for label, code in filters:
            lst.add_widget(OneLineListItem(text=label, on_release=lambda x, c=code: self._apply_history_filter(c)))
        scroll.add_widget(lst)
        self.hist_filter_dialog = MDDialog(title='Filtrer par type', type='custom', content_cls=scroll, size_hint=(0.85, None))
        self.hist_filter_dialog.open()

    def _apply_history_filter(self, code: str):
        self.current_history_filter = code
        if getattr(self, 'hist_filter_dialog', None):
            self.hist_filter_dialog.dismiss()
        if hasattr(self, 'btn_hist_filter'):
            self.btn_hist_filter.md_bg_color = (0.2, 0.2, 0.2, 1) if code == AppConstants.FAM_ALL else (0, 0.55, 0.8, 1)
            self.btn_hist_filter.text = "FILTRER PAR TYPE D'OPÉRATION" if code == AppConstants.FAM_ALL else f'FILTRE : {code}'
        self.history_page_offset = 0
        self.history_has_more = True
        self.rv_history.data = []
        self.load_more_history()

    def open_history_date_picker(self, instance):
        dp_dlg = MDDatePicker(mode='range')
        dp_dlg.bind(on_save=lambda inst, val, dr: self.filter_history_list('custom', dr[0], dr[-1]) if dr else self.filter_history_list('custom', val, val))
        dp_dlg.open()

    def filter_history_list(self, ftype: str = 'auj', start_d=None, end_d=None):
        inact, act = (0.5, 0.5, 0.5, 1), self.theme_cls.primary_color
        for b in (self.btn_hist_today, self.btn_hist_yest, self.btn_hist_week, self.btn_hist_month, self.btn_hist_date):
            b.md_bg_color = inact
        today = datetime.now().date()
        if ftype == 'custom' and start_d:
            self.history_start_date = start_d
            self.history_end_date = end_d or start_d
            self.btn_hist_date.md_bg_color = act
            self.btn_hist_date.text = start_d.strftime('%d/%m') if start_d == self.history_end_date else 'CAL'
        else:
            self.btn_hist_date.text = 'CAL'
            if ftype == 'auj':
                self.history_start_date = self.history_end_date = today
                self.btn_hist_today.md_bg_color = act
            elif ftype == 'hier':
                self.history_start_date = self.history_end_date = today - timedelta(days=1)
                self.btn_hist_yest.md_bg_color = act
            elif ftype == 'semaine':
                self.history_start_date = today - timedelta(days=(today.weekday() + 2) % 7)
                self.history_end_date = today
                self.btn_hist_week.md_bg_color = act
            elif ftype == 'mois':
                self.history_start_date = today.replace(day=1)
                self.history_end_date = today
                self.btn_hist_month.md_bg_color = act

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
        try:
            rows = self.db.get_operations_history(
                str(self.history_start_date),
                str(self.history_end_date),
                doc_type_filter=self.current_history_filter,
                limit=self.history_batch_size,
                offset=self.history_page_offset,
            )
            self.history_has_more = len(rows) == self.history_batch_size
            formatted = self._format_history_rows(rows)
            if self.history_page_offset == 0:
                if not formatted:
                    formatted = [{
                        'type_str': 'Information', 'ref_str': '', 'entity_str': 'Aucune opération trouvée.',
                        'date_str': '', 'amount_text': '', 'icon': 'information-outline',
                        'icon_color': [0.5, 0.5, 0.5, 1], 'bg_color': [1, 1, 1, 1],
                        'op_id': None, 'is_transfer': False,
                    }]
                self.rv_history.data = formatted
            else:
                self.rv_history.data.extend(formatted)
            self.history_page_offset += len(rows)
            self.rv_history.refresh_from_data()
        finally:
            if hasattr(self, 'rv_history'):
                self.rv_history.loading_lock = False
            self.is_loading_history = False

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
        vis = AppConstants.DOC_VISUALS.get(dt_code, {'name': dt_code})

        ent_cat = tx.get('entity_category', AppConstants.get_entity_type(dt_code))
        is_supp = ent_cat == AppConstants.ENTITY_SUPPLIER
        is_comptoir = AppConstants.is_default_entity(tx.get('entity_id'), ent_cat)
        total_amt = FinancialEngine.to_decimal(tx.get('total_amount', 0))
        paid_amt = FinancialEngine.to_decimal(tx.get('paid_amount', 0))
        timbre_val = FinancialEngine.to_decimal(tx.get('stamp_duty', 0))

        content = MDBoxLayout(orientation='vertical', spacing=10, size_hint_y=None, height=dp(550))
        hbox = MDCard(orientation='vertical', adaptive_height=True, padding=dp(12), md_bg_color=(0.95, 0.96, 0.98, 1), radius=[10], elevation=0)
        hbox.add_widget(MDLabel(text=self.fix_text(f"{vis['name']} ({tx.get('doc_number', '')})"), bold=True, font_style='Subtitle1', adaptive_height=True))
        if is_transfer:
            src = AppConstants.LOC_LABELS.get(tx.get('source_location'), 'Magasin')
            dst = AppConstants.LOC_LABELS.get(tx.get('destination_location'), 'Dépôt')
            hbox.add_widget(MDLabel(text=f'Trajet : {src} -> {dst}', font_style='Subtitle2', adaptive_height=True))
        else:
            lbl_e = 'Fournisseur' if is_supp else 'Client'
            hbox.add_widget(MDLabel(text=f"{lbl_e} : {self.fix_text(ent.get('name', AppConstants.DEFAULT_CLIENT_NAME))}", font_style='Subtitle2', adaptive_height=True))
        hbox.add_widget(MDLabel(text=f"Date : {str(tx.get('date', ''))[:16]}", font_style='Caption', theme_text_color='Secondary', adaptive_height=True))

        if not is_transfer and dt_code not in AppConstants.QTY_ONLY_DOC_TYPES:
            if timbre_val > 0:
                hbox.add_widget(MDLabel(text=f'Droit de Timbre : {FinancialEngine.format_money(timbre_val)}', theme_text_color='Custom', text_color=(0.5, 0, 0.5, 1), font_style='Caption', adaptive_height=True))
            hbox.add_widget(MDLabel(text=f'Montant : {FinancialEngine.format_money(abs(total_amt))}', bold=True, font_style='H5', adaptive_height=True))
            if AppConstants.FINANCIAL_FACTORS.get(dt_code, 0) == 1 and not is_comptoir:
                rem = FinancialEngine.quantize_money(total_amt - paid_amt)
                if abs(rem) < Decimal('0.05'):
                    hbox.add_widget(MDLabel(text='Statut : Payée', theme_text_color='Custom', text_color=(0, 0.65, 0.2, 1), bold=True, adaptive_height=True))
                else:
                    hbox.add_widget(MDLabel(text=f"{'Règlement' if is_supp else 'Versement'} : {FinancialEngine.format_money(paid_amt)}", theme_text_color='Custom', text_color=(0, 0.6, 0.2, 1), bold=True, adaptive_height=True))
                    hbox.add_widget(MDLabel(text=f'Reste : {FinancialEngine.format_money(rem)}', theme_text_color='Custom', text_color=(0.85, 0, 0, 1), bold=True, adaptive_height=True))
        content.add_widget(hbox)

        scroll = MDScrollView()
        lst = MDList()
        if items:
            for it in items:
                q_disp = FinancialEngine.format_qty(it['qty'])
                ib = MDBoxLayout(orientation='vertical', adaptive_height=True, padding=[dp(10), dp(8)], spacing=dp(2))
                ib.add_widget(MDLabel(text=self.fix_text(it['product_name']), font_style='Subtitle1', bold=True, adaptive_height=True))
                if is_transfer or dt_code in AppConstants.QTY_ONLY_DOC_TYPES:
                    ib.add_widget(MDLabel(text=f'Qté : {q_disp}', theme_text_color='Custom', text_color=(0, 0.4, 0.8, 1), bold=True, adaptive_height=True))
                else:
                    lc = FinancialEngine.calculate_document_totals([it], doc_type=dt_code)
                    p_disp = FinancialEngine.format_money(it['price'])
                    ib.add_widget(MDLabel(text=f"{q_disp} x {p_disp} = {FinancialEngine.format_money(lc['final_total'])}", theme_text_color='Secondary', font_style='Body2', adaptive_height=True))
                lst.add_widget(ib)
                lst.add_widget(MDSeparator(height=dp(1), color=(0.9, 0.9, 0.9, 1)))
        else:
            lst.add_widget(OneLineListItem(text=f"Note : {self.fix_text(tx.get('note', 'Opération financière'))}"))
        scroll.add_widget(lst)
        content.add_widget(scroll)

        actions = MDBoxLayout(orientation='vertical', spacing='10dp', adaptive_height=True, padding=[0, '10dp', 0, 0])
        top_row = MDBoxLayout(orientation='horizontal', spacing='10dp', size_hint_y=None, height='48dp')
        if dt_code not in AppConstants.INVOICE_TAX_DOC_TYPES:
            top_row.add_widget(MDFillRoundFlatButton(text='IMPRIMER', md_bg_color=(0, 0.5, 0.8, 1), size_hint_x=1, on_release=lambda x: threading.Thread(target=self.print_ticket_bluetooth, args=(details,), daemon=True).start()))
        if dt_code not in (AppConstants.DOC_CLIENT_PAY, AppConstants.DOC_SUPPLIER_PAY):
            top_row.add_widget(MDFillRoundFlatButton(text='PDF A4', md_bg_color=(0.8, 0.2, 0.2, 1), size_hint_x=1, on_release=lambda x: self.generate_pdf_report(op_id, is_transfer)))

        is_today = str(tx.get('date', '')).split(' ')[0] == datetime.now().strftime('%Y-%m-%d')
        if (not self.is_seller_mode or is_today) and dt_code != AppConstants.DOC_BI:
            top_row.add_widget(MDFillRoundFlatButton(text='MODIFIER', md_bg_color=(0, 0.6, 0.4, 1), size_hint_x=1, on_release=lambda x: self.load_transaction_for_edit(details)))
        actions.add_widget(top_row)

        if not self.is_seller_mode or is_today:
            actions.add_widget(MDFlatButton(text='SUPPRIMER', theme_text_color='Custom', text_color=(0.9, 0, 0, 1), size_hint_x=1, on_release=lambda x: self.confirm_delete_transaction(op_id, is_transfer)))
        content.add_widget(actions)

        self.srv_dialog = MDDialog(title='Détails', type='custom', content_cls=content, size_hint=(0.95, 0.95), buttons=[MDFlatButton(text='FERMER', on_release=lambda x: self.srv_dialog.dismiss())])
        self.srv_dialog.open()

    def load_transaction_for_edit(self, details: dict):
        if getattr(self, 'srv_dialog', None):
            self.srv_dialog.dismiss()
        if getattr(self, 'pending_dialog', None):
            self.pending_dialog.dismiss()

        tx = details['transaction']
        ent = details['entity']
        items = details['items']
        dt_code = tx['doc_type']

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
        def _do_del(x):
            if getattr(self, 'confirm_del_dialog', None):
                self.confirm_del_dialog.dismiss()
            if getattr(self, 'srv_dialog', None):
                self.srv_dialog.dismiss()
            try:
                self.db.delete_transaction(op_id, is_transfer=is_transfer)
                self.notify('Opération supprimée.', 'success')
                if getattr(self, 'pending_dialog', None):
                    self.filter_history_list('auj')
                if getattr(self, 'entity_hist_dialog', None):
                    self.filter_entity_history_list('auj')
                self.load_more_entities(reset=True)
                self.check_and_load_stats()
            except Exception as exc:
                self.notify(str(exc), 'error')

        self.confirm_del_dialog = MDDialog(
            title='Confirmation',
            text='Voulez-vous vraiment supprimer cette opération ?\nLe stock et le solde seront recalculés automatiquement.',
            buttons=[
                MDFlatButton(text='NON', on_release=lambda x: self.confirm_del_dialog.dismiss()),
                MDRaisedButton(text='OUI, SUPPRIMER', md_bg_color=(0.8, 0, 0, 1), on_release=_do_del),
            ],
        )
        self.confirm_del_dialog.open()

    def generate_pdf_report(self, op_id: int, is_transfer: bool = False):
        now = time.time()
        if getattr(self, '_is_generating_pdf', False) or (now - getattr(self, '_last_pdf_click_time', 0) < 1.5):
            return
        self._last_pdf_click_time = now
        self._is_generating_pdf = True
        temp_pdf = os.path.join(self.user_data_dir, f'temp_{int(time.time())}.pdf')
        try:
            details = self.db.get_operation_details(op_id, is_transfer=is_transfer)
            if not details:
                self.notify('Opération introuvable.', 'error')
                return
            tx = details['transaction']
            ent = details['entity']
            items = details['items']
            dt_code = tx['doc_type']
            store_info = self.db.get_store_info()

            pdf = PDF(orientation='P', unit='mm', format='A4')
            pdf.store_info = store_info
            ent_cat = tx.get('entity_category', AppConstants.get_entity_type(dt_code))
            is_supp = ent_cat == AppConstants.ENTITY_SUPPLIER
            is_comptoir = AppConstants.is_default_entity(tx.get('entity_id'), ent_cat)

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
                }

            vis = AppConstants.DOC_VISUALS.get(dt_code, {'name': dt_code})
            try:
                date_str = datetime.strptime(str(tx['date'])[:19], '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y')
            except Exception:
                date_str = str(tx.get('date', ''))[:10]

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
                    cur_bal = FinancialEngine.to_decimal(ent.get('balance', 0))
                    eff = calc['final_total'] - FinancialEngine.to_decimal(tx.get('paid_amount', 0))
                    pdf.balance_data = {'old_balance': cur_bal - eff, 'transaction_amount': eff, 'new_balance': cur_bal}

            pdf.add_page()
            pdf.draw_table_with_fill(headers, table_data, col_widths)

            safe_client = ''.join(c for c in disp_ent if c.isalnum() or c in (' ', '-', '_')).strip().replace(' ', '_') or 'Client'
            safe_num = str(doc_num).replace('/', '-').replace('\\', '-').replace(':', '')
            fname = f"{vis['name'].replace(' ', '_')}_{safe_num}_{safe_client}.pdf"

            pdf.output(temp_pdf)
            with open(temp_pdf, 'rb') as f:
                pdf_bytes = f.read()
            saved_path, saved_uri = self._save_bytes_to_downloads(fname, pdf_bytes, 'application/pdf')
            self.notify(f'PDF enregistré : {fname}', 'success')
            self._show_pdf_ready_dialog(fname, saved_path, saved_uri)
        except Exception as exc:
            self.notify(f'Erreur PDF : {exc}', 'error')
        finally:
            self._is_generating_pdf = False
            if os.path.exists(temp_pdf):
                try:
                    os.remove(temp_pdf)
                except Exception:
                    pass

    # --------------------------------------------------------------------------
    # 7.11 IMPRESSION BLUETOOTH ESC/POS 80 MM
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
            content = MDBoxLayout(orientation='vertical', size_hint_y=None, height=dp(400))
            scroll = MDScrollView()
            lst = MDList()
            if not paired:
                lst.add_widget(OneLineListItem(text='Aucun appareil associé (Paired)'))
            else:
                for dev in paired:
                    dn, dm = dev.getName(), dev.getAddress()
                    it = TwoLineAvatarIconListItem(text=dn, secondary_text=dm, on_release=lambda x, n=dn, m=dm: self.select_printer(n, m))
                    it.add_widget(IconLeftWidget(icon='printer-wireless'))
                    lst.add_widget(it)
            scroll.add_widget(lst)
            content.add_widget(scroll)
            self.bt_dialog = MDDialog(title='Choisir Imprimante', type='custom', content_cls=content, buttons=[MDFlatButton(text='ANNULER', on_release=lambda x: self.bt_dialog.dismiss())])
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

            try:
                f_modes = params.getSupportedFocusModes()
                if f_modes is not None:
                    modes_list = [str(m) for m in f_modes.toArray()]
                    if 'continuous-picture' in modes_list:
                        params.setFocusMode('continuous-picture')
                    elif 'continuous-video' in modes_list:
                        params.setFocusMode('continuous-video')
                    elif 'auto' in modes_list:
                        params.setFocusMode('auto')
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
                        default_idx = 0
                        for idx, r_val in enumerate(self._zoom_ratios):
                            if r_val >= 100:
                                default_idx = idx
                                break
                        try:
                            params.setZoom(default_idx)
                            self._current_zoom = (self._zoom_ratios[default_idx] / 100.0) if self._zoom_ratios else 1.0
                        except Exception:
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

    def _launch_camera_widget(self):
        from kivy.graphics import Color, Line, Rectangle
        self.scanner_start_time = time.time()
        self._laser_phase = 0.0
        self._laser_dir = 1.0
        try:
            try:
                self.camera_widget = Camera(play=True, index=0, resolution=(1280, 720), allow_stretch=True, keep_ratio=False)
            except Exception:
                self.camera_widget = Camera(play=True, index=0, resolution=(640, 480), allow_stretch=True, keep_ratio=False)
            with self.camera_widget.canvas.before:
                PushMatrix()
                self.rotation = Rotate(angle=-90 if platform == 'android' else 0, origin=self.camera_widget.center)
            with self.camera_widget.canvas.after:
                PopMatrix()
            self.camera_widget.bind(center=lambda i, v: setattr(self.rotation, 'origin', i.center))
        except Exception:
            self.notify('Erreur chargement caméra.', 'error')
            return

        is_single = getattr(self, 'target_scan_field', None) is not None
        root = MDBoxLayout(orientation='vertical')
        cam_area = MDFloatLayout(size_hint_y=0.90 if is_single else 0.48)
        self.cam_area_ref = cam_area
        self.camera_widget.size_hint = (1, 1)
        cam_area.add_widget(self.camera_widget)

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
        Clock.schedule_once(lambda dt: self._update_scan_overlay_geometry(), 0.05)
        Clock.schedule_once(lambda dt: self._optimize_camera_params(), 0.35)
        Clock.schedule_once(lambda dt: self._optimize_camera_params(), 0.90)
        self._laser_ev = Clock.schedule_interval(self._animate_laser_step, 1.0 / 30.0)
        self.scan_event = Clock.schedule_interval(self.detect_barcode_frame, 1.0 / 20.0)

    def _update_scan_overlay_geometry(self, *args):
        cam_area = getattr(self, 'cam_area_ref', None)
        if not cam_area or not getattr(self, '_scan_box_line', None):
            return
        x, y = cam_area.pos
        w, h = cam_area.size
        box_w = max(dp(40), w * 0.84)
        box_h = max(dp(40), h * 0.56)
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
            self.camera_widget.play = False
        if getattr(self, 'scan_dialog', None):
            self.scan_dialog.dismiss()
            self.scan_dialog = None
        self.cam_area_ref = None
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
            threading.Thread(
                target=self._decode_frame_worker,
                args=(raw_pixels, tex_size, color_fmt, sid),
                daemon=True,
            ).start()
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
    # 7.13 PARAMÈTRES, MAGASIN, SÉCURITÉ VENDEUR, EXCEL & SAUVEGARDE ZIP
    # --------------------------------------------------------------------------
    def open_settings_menu(self, *args):
        if getattr(self, 'settings_menu_dialog', None):
            try:
                self.settings_menu_dialog.dismiss()
            except Exception:
                pass
            self.settings_menu_dialog = None

        scroll_height = min(dp(340) if self.is_seller_mode else dp(500), Window.height * 0.62)
        scroll_view = MDScrollView(size_hint_y=None, height=scroll_height, bar_width=dp(3))
        content = MDBoxLayout(
            orientation='vertical',
            spacing=dp(12),
            size_hint_y=None,
            adaptive_height=True,
            padding=[dp(2), dp(8), dp(4), dp(8)],
        )

        def _add_card_header(parent_card, icon_name: str, title_text: str):
            hdr = MDBoxLayout(orientation='horizontal', spacing=dp(8), size_hint_y=None, height=dp(24))
            hdr.add_widget(MDIcon(icon=icon_name, theme_text_color='Custom', text_color=self.theme_cls.primary_color, font_size='20sp', pos_hint={'center_y': 0.5}))
            hdr.add_widget(MDLabel(text=title_text, font_style='Caption', bold=True, theme_text_color='Custom', text_color=self.theme_cls.primary_color, pos_hint={'center_y': 0.5}))
            parent_card.add_widget(hdr)

        def _add_action_row(parent_card, title_txt: str, sub_txt: str, icon_name: str, icon_col, bg_col, cb, right_icon='chevron-right', title_col=None):
            row_card = MDCard(
                orientation='horizontal',
                padding=[dp(12), dp(8), dp(12), dp(8)],
                spacing=dp(10),
                size_hint_y=None,
                height=dp(56),
                radius=[10],
                elevation=0,
                md_bg_color=bg_col,
                ripple_behavior=True,
                on_release=cb,
            )
            row_card.add_widget(MDIcon(icon=icon_name, theme_text_color='Custom', text_color=icon_col, font_size='24sp', pos_hint={'center_y': 0.5}))
            t_box = MDBoxLayout(orientation='vertical', pos_hint={'center_y': 0.5}, spacing=dp(2))
            if title_col:
                t_box.add_widget(MDLabel(text=title_txt, bold=True, font_style='Subtitle2', theme_text_color='Custom', text_color=title_col, shorten=True))
            else:
                t_box.add_widget(MDLabel(text=title_txt, bold=True, font_style='Subtitle2', theme_text_color='Primary', shorten=True))
            t_box.add_widget(MDLabel(text=sub_txt, font_style='Caption', theme_text_color='Secondary', shorten=True))
            row_card.add_widget(t_box)
            row_card.add_widget(MDIcon(icon=right_icon, theme_text_color='Hint', font_size='22sp', pos_hint={'center_y': 0.5}))
            parent_card.add_widget(row_card)

        def _add_checkbox_row(parent_card, label_txt: str, setting_key: str, default_val: bool = False):
            row = MDBoxLayout(orientation='horizontal', size_hint_y=None, adaptive_height=True, padding=[dp(4), dp(2), 0, dp(2)], spacing=dp(8))
            lbl = MDLabel(text=label_txt, theme_text_color='Primary', font_style='Body2', size_hint_x=1, size_hint_y=None, pos_hint={'center_y': 0.5})
            lbl.bind(width=lambda inst, val: setattr(inst, 'text_size', (val, None)), texture_size=lambda inst, val: setattr(inst, 'height', max(dp(24), val[1])))
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
                md_bg_color=(0.94, 0.98, 0.94, 1),
                radius=[12],
                elevation=1,
                ripple_behavior=True,
                on_release=lambda inst: [Clipboard.copy(activ_key), self.notify("Clé d'activation copiée", 'success')],
            )
            top_row = MDBoxLayout(orientation='horizontal', adaptive_height=True, spacing=dp(10))
            top_row.add_widget(MDIcon(icon='shield-check', theme_text_color='Custom', text_color=(0, 0.65, 0.2, 1), font_size='26sp', pos_hint={'center_y': 0.5}))
            top_row.add_widget(MDLabel(text='Licence Activée', font_style='Subtitle1', bold=True, theme_text_color='Custom', text_color=(0, 0.65, 0.2, 1), pos_hint={'center_y': 0.5}))
            lic_card.add_widget(top_row)
            lic_card.add_widget(MDLabel(text="Clé d'activation (appuyez pour copier) :", font_style='Caption', theme_text_color='Secondary', adaptive_height=True))
            lbl_key = MDLabel(text=activ_key, font_style='Caption', theme_text_color='Primary', font_name='Roboto', bold=True, size_hint_y=None)
            lbl_key.bind(width=lambda inst, val: setattr(inst, 'text_size', (val, None)), texture_size=lambda inst, val: setattr(inst, 'height', max(dp(18), val[1])))
            lic_card.add_widget(lbl_key)
        else:
            lic_card = MDCard(
                orientation='vertical',
                padding=dp(14),
                spacing=dp(6),
                size_hint_y=None,
                adaptive_height=True,
                md_bg_color=(1, 0.95, 0.9, 1),
                radius=[12],
                elevation=1,
                ripple_behavior=True,
                on_release=lambda inst: [Clipboard.copy(dev_id), self.notify('ID copié', 'success')],
            )
            top_row = MDBoxLayout(orientation='horizontal', adaptive_height=True, spacing=dp(10))
            top_row.add_widget(MDIcon(icon='shield-alert', theme_text_color='Custom', text_color=(0.9, 0.5, 0, 1), font_size='26sp', pos_hint={'center_y': 0.5}))
            top_row.add_widget(MDLabel(text=f'Essai gratuit : {days_left} jour(s) restant(s)', font_style='Subtitle1', bold=True, theme_text_color='Custom', text_color=(0.9, 0.5, 0, 1), pos_hint={'center_y': 0.5}))
            lic_card.add_widget(top_row)
            lic_card.add_widget(MDLabel(text="ID de l'appareil (appuyez pour copier) :", font_style='Caption', theme_text_color='Secondary', adaptive_height=True))
            lic_card.add_widget(MDLabel(text=dev_id, font_style='Caption', theme_text_color='Primary', font_name='Roboto', bold=True, adaptive_height=True))
        content.add_widget(lic_card)

        if not self.is_seller_mode:
            # 2. Carte Magasin & Règles
            store_card = MDCard(orientation='vertical', padding=dp(12), spacing=dp(8), size_hint_y=None, adaptive_height=True, radius=[12], elevation=1, md_bg_color=(0.97, 0.97, 0.99, 1))
            _add_card_header(store_card, 'store-cog', 'MAGASIN & RÈGLES')
            store_name = self.db.get_setting('store_name', '') or AppConstants.APP_NAME
            _add_action_row(store_card, 'Informations Magasin', self.fix_text(f'Actuel : {store_name}'), 'store', (0.1, 0.45, 0.8, 1), (0.93, 0.96, 1, 1), lambda x: self.show_store_settings_dialog())
            store_card.add_widget(MDSeparator(height=dp(1), color=(0.9, 0.9, 0.92, 1)))
            _add_checkbox_row(store_card, 'Autoriser la vente en stock négatif', 'allow_negative_stock', False)
            store_card.add_widget(MDSeparator(height=dp(1), color=(0.9, 0.9, 0.92, 1)))
            _add_checkbox_row(store_card, 'Avertir si prix vente < prix achat', 'warn_price_below_cost', True)
            content.add_widget(store_card)

            # 3. Carte Données & Sauvegarde
            data_card = MDCard(orientation='vertical', padding=dp(12), spacing=dp(8), size_hint_y=None, adaptive_height=True, radius=[12], elevation=1, md_bg_color=(0.97, 0.97, 0.99, 1))
            _add_card_header(data_card, 'database-cog', 'DONNÉES & SAUVEGARDE')
            _add_action_row(data_card, 'Exporter Produits', 'Direct Excel (.xlsx)', 'file-export', (0.1, 0.6, 0.3, 1), (1, 1, 1, 1), lambda x: self.perform_export())
            _add_action_row(data_card, 'Importer Produits', 'Depuis Excel (.xlsx)', 'file-import', (0.1, 0.45, 0.8, 1), (1, 1, 1, 1), lambda x: [self.settings_menu_dialog.dismiss(), self.import_data_dialog()])
            _add_action_row(data_card, 'Sauvegarde Locale', 'Archive complète (.zip)', 'database-export', (0.4, 0.2, 0.7, 1), (1, 1, 1, 1), lambda x: self.perform_local_backup('manual'))
            _add_action_row(data_card, 'Partager Sauvegarde', 'Quick Share & Bluetooth (.zip)', 'cloud-upload', (0.0, 0.55, 0.65, 1), (1, 1, 1, 1), lambda x: self.share_database_file())
            _add_action_row(data_card, 'Restaurer', 'Depuis une archive (.zip / .db)', 'backup-restore', (0.85, 0.45, 0.0, 1), (1, 0.97, 0.92, 1), lambda x: [self.settings_menu_dialog.dismiss(), self.show_restore_dialog()])
            content.add_widget(data_card)

            # 4. Carte Imprimante & PDF
            printer_card = MDCard(orientation='vertical', padding=dp(12), spacing=dp(8), size_hint_y=None, adaptive_height=True, radius=[12], elevation=1, md_bg_color=(0.97, 0.97, 0.99, 1))
            _add_card_header(printer_card, 'printer', 'IMPRIMANTE & PDF')
            p_name = self.db.get_setting('printer_name', 'Non configurée') or 'Non configurée'
            _add_action_row(printer_card, 'Configurer Imprimante', f'Actuelle : {p_name}', 'printer-wireless', (0.1, 0.45, 0.8, 1), (1, 1, 1, 1), lambda x: [self.settings_menu_dialog.dismiss(), self.open_bluetooth_selector(x)])
            _add_action_row(printer_card, "Oublier l'imprimante", "Dissocier l'appareil actuel", 'printer-off', (0.8, 0.15, 0.15, 1), (1, 0.95, 0.95, 1), self.clear_printer_selection, title_col=(0.8, 0.15, 0.15, 1))
            printer_card.add_widget(MDSeparator(height=dp(1), color=(0.9, 0.9, 0.92, 1)))
            _add_checkbox_row(printer_card, 'Impression Auto (Bluetooth 80 mm)', 'printer_auto', False)
            printer_card.add_widget(MDSeparator(height=dp(1), color=(0.9, 0.9, 0.92, 1)))
            _add_checkbox_row(printer_card, 'Afficher le solde sur PDF (BV / BA)', 'show_balance_in_pdf', False)
            content.add_widget(printer_card)

        # 5. Carte Sécurité
        sec_card = MDCard(orientation='vertical', padding=dp(12), spacing=dp(8), size_hint_y=None, adaptive_height=True, radius=[12], elevation=1, md_bg_color=(0.97, 0.97, 0.99, 1))
        _add_card_header(sec_card, 'shield-account', 'SÉCURITÉ')
        mode_lbl = 'Mode Vendeur (Restreint)' if self.is_seller_mode else 'Mode Administrateur (Complet)'
        _add_action_row(
            sec_card,
            'Mode Admin / Vendeur',
            f'Actif : {mode_lbl}',
            'shield-key',
            (0.3, 0.3, 0.75, 1),
            (1, 1, 1, 1),
            lambda x: [self.settings_menu_dialog.dismiss(), Clock.schedule_once(lambda dt: self.open_seller_auth_dialog(), 0.12)],
        )
        content.add_widget(sec_card)

        # 6. Carte Mise à jour & Version
        update_card = MDCard(
            orientation='horizontal',
            padding=[dp(14), dp(10), dp(14), dp(10)],
            spacing=dp(12),
            size_hint_y=None,
            height=dp(60),
            radius=[12],
            elevation=1,
            md_bg_color=(0.92, 0.97, 0.97, 1),
            ripple_behavior=True,
            on_release=lambda x: [self.settings_menu_dialog.dismiss(), webbrowser.open('https://mobile.magpro-soft.com/')],
        )
        update_card.add_widget(MDIcon(icon='cloud-download', theme_text_color='Custom', text_color=(0.1, 0.55, 0.55, 1), font_size='28sp', pos_hint={'center_y': 0.5}))
        upd_txt = MDBoxLayout(orientation='vertical', pos_hint={'center_y': 0.5}, spacing=dp(2))
        upd_txt.add_widget(MDLabel(text="Mise à jour de l'application", bold=True, font_style='Subtitle1', theme_text_color='Primary'))
        upd_txt.add_widget(MDLabel(text=f'Version actuelle : {AppConstants.APP_VERSION} • © {AppConstants.COPYRIGHT_YEAR}', font_style='Caption', theme_text_color='Secondary'))
        update_card.add_widget(upd_txt)
        update_card.add_widget(MDIcon(icon='open-in-new', theme_text_color='Hint', font_size='22sp', pos_hint={'center_y': 0.5}))
        content.add_widget(update_card)

        scroll_view.add_widget(content)
        self.settings_menu_dialog = MDDialog(
            title='Paramètres & Configuration',
            type='custom',
            content_cls=scroll_view,
            radius=[15, 15, 15, 15],
            buttons=[MDFlatButton(text='FERMER', theme_text_color='Error', on_release=lambda x: self.settings_menu_dialog.dismiss())],
            size_hint=(0.95, None),
        )
        self.settings_menu_dialog.open()

    def show_store_settings_dialog(self):
        if getattr(self, 'settings_menu_dialog', None):
            self.settings_menu_dialog.dismiss()
        scroll = MDScrollView(size_hint_y=None, height=dp(560))
        box = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(12), padding=dp(10))

        fields = {}
        for key, label in (
            ('store_name', 'Nom du Magasin'),
            ('store_activity', 'Activité'),
            ('store_phone', 'Téléphone'),
            ('store_address', 'Adresse'),
            ('store_email', 'Email'),
            ('store_rc', 'N° RC'),
            ('store_nif', 'N.I.F'),
            ('store_nis', 'N.I.S'),
            ('store_nai', 'N.A.I'),
        ):
            tf = SmartTextField(text=self.db.get_setting(key, ''), hint_text=label)
            fields[key] = tf
            box.add_widget(tf)

        def _save(x):
            for k, tf in fields.items():
                self.db.set_setting(k, tf.get_value().strip())
            self.notify('Informations du magasin enregistrées.', 'success')
            self.store_settings_dialog.dismiss()
            self.open_settings_menu()

        box.add_widget(MDRaisedButton(text='ENREGISTRER', md_bg_color=(0, 0.7, 0.2, 1), size_hint_x=1, height=dp(48), on_release=_save))
        box.add_widget(MDRaisedButton(text='ANNULER', md_bg_color=(0.9, 0.9, 0.9, 1), text_color=(0.2, 0.2, 0.2, 1), size_hint_x=1, height=dp(48), on_release=lambda x: [self.store_settings_dialog.dismiss(), self.open_settings_menu()]))
        scroll.add_widget(box)
        self.store_settings_dialog = MDDialog(title='Informations du Magasin', type='custom', content_cls=scroll, size_hint=(0.95, 0.92))
        self.store_settings_dialog.open()

    def open_seller_auth_dialog(self):
        if getattr(self, 'auth_dialog', None):
            try:
                self.auth_dialog.dismiss()
            except Exception:
                pass
            self.auth_dialog = None
        Window.release_all_keyboards()

        stored = self.db.get_setting('seller_password', '') or ''
        has_pass = bool(stored)
        self.temp_new_seller_pass = None
        content = MDBoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None, height=dp(80))
        self.seller_pass_field = MDTextField(hint_text='Code PIN' if has_pass else 'Créer un code PIN', password=True, halign='center')
        content.add_widget(self.seller_pass_field)
        btn = MDRaisedButton(
            text='VALIDER' if has_pass else 'SUIVANT',
            md_bg_color=(0, 0.55, 0.8, 1),
            on_release=lambda x: self._verify_seller_pin(has_pass, stored, btn),
        )
        self.seller_pass_field.bind(on_text_validate=lambda x: self._verify_seller_pin(has_pass, stored, btn))
        self.auth_dialog = MDDialog(
            title='Accès Admin' if has_pass else 'Nouveau Code PIN',
            type='custom',
            content_cls=content,
            size_hint=(0.88, None),
            buttons=[MDFlatButton(text='ANNULER', on_release=lambda x: self.auth_dialog.dismiss()), btn],
        )
        self.auth_dialog.open()

    def _verify_seller_pin(self, has_pass: bool, stored: str, btn):
        pwd = self.seller_pass_field.text.strip()
        if not pwd:
            self.notify('Le code PIN est vide.', 'error')
            return
        if has_pass:
            if pwd == stored:
                self.seller_pass_field.focus = False
                Window.release_all_keyboards()
                self.auth_dialog.dismiss()
                Clock.schedule_once(lambda dt: self._open_seller_toggle_dialog(), 0.12)
            else:
                self.notify('Code PIN incorrect.', 'error')
                self.seller_pass_field.text = ''
        elif self.temp_new_seller_pass is None:
            self.temp_new_seller_pass = pwd
            self.auth_dialog.title = 'Confirmer le code PIN'
            self.seller_pass_field.text = ''
            btn.text = 'CONFIRMER'
        elif pwd == self.temp_new_seller_pass:
            self.db.set_setting('seller_password', pwd)
            self.seller_pass_field.focus = False
            Window.release_all_keyboards()
            self.auth_dialog.dismiss()
            Clock.schedule_once(lambda dt: self._open_seller_toggle_dialog(), 0.12)
        else:
            self.notify('Les codes ne correspondent pas.', 'error')
            self.temp_new_seller_pass = None
            self.auth_dialog.title = 'Nouveau Code PIN'
            self.seller_pass_field.text = ''
            btn.text = 'SUIVANT'

    def _open_seller_toggle_dialog(self):
        if getattr(self, 'toggle_dialog', None):
            try:
                self.toggle_dialog.dismiss()
            except Exception:
                pass
            self.toggle_dialog = None

        content = MDBoxLayout(
            orientation='horizontal',
            spacing=dp(12),
            size_hint_y=None,
            height=dp(56),
            padding=[dp(16), 0, dp(8), 0],
        )
        lbl = MDLabel(
            text='Mode Vendeur (Restreint)',
            font_style='Subtitle1',
            theme_text_color='Primary',
            size_hint_x=1,
            pos_hint={'center_y': 0.5},
        )
        chk = MDCheckbox(
            active=bool(self.is_seller_mode),
            size_hint=(None, None),
            size=(dp(48), dp(48)),
            pos_hint={'center_y': 0.5},
        )
        chk.bind(active=self._on_seller_mode_checkbox_active)

        def _on_label_touch(inst, touch):
            if inst.collide_point(*touch.pos) and not chk.collide_point(*touch.pos):
                chk.active = not chk.active
                return True
            return False

        lbl.bind(on_touch_down=_on_label_touch)
        content.add_widget(lbl)
        content.add_widget(chk)

        self.toggle_dialog = MDDialog(
            title='Configuration du Mode',
            type='custom',
            content_cls=content,
            size_hint=(0.9, None),
            buttons=[
                MDFlatButton(
                    text='FERMER',
                    on_release=lambda x: self.toggle_dialog.dismiss() if getattr(self, 'toggle_dialog', None) else None,
                )
            ],
        )
        self.toggle_dialog.open()

    def _on_seller_mode_checkbox_active(self, instance, val: bool):
        val = bool(val)
        if self.is_seller_mode == val:
            return
        self.is_seller_mode = val
        self.db.set_setting('config_seller_mode', str(val))
        self.notify(f"Mode Vendeur : {'Activé' if val else 'Désactivé'}", 'info')
        if getattr(self, '_seller_layout_ev', None):
            self._seller_layout_ev.cancel()
        self._seller_layout_ev = Clock.schedule_once(lambda dt: self.update_dashboard_layout(), 0.05)

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
                cast('android.app.Activity', PythonActivity.mActivity).startActivity(
                    intent_cls.createChooser(intent, str_cls(f'Partager : {os.path.basename(saved_path)}'))
                )
        except Exception as exc:
            self.notify(f'Erreur partage/ouverture : {exc}', 'error')

    def _show_pdf_ready_dialog(self, fname: str, saved_path: str, saved_uri):
        if getattr(self, 'pdf_ready_dialog', None):
            try:
                self.pdf_ready_dialog.dismiss()
            except Exception:
                pass

        def _on_pdf_action(btn_inst, act_mode: str):
            if getattr(btn_inst, '_clicked', False):
                return
            btn_inst._clicked = True
            btn_inst.disabled = True
            if getattr(self, 'pdf_ready_dialog', None):
                self.pdf_ready_dialog.dismiss()
            self._share_or_open_exported_file(saved_path, saved_uri, 'application/pdf', action=act_mode)

        btn_close = MDFlatButton(
            text='FERMER',
            on_release=lambda x: self.pdf_ready_dialog.dismiss() if getattr(self, 'pdf_ready_dialog', None) else None,
        )
        btn_open = MDRaisedButton(
            text='OUVRIR',
            md_bg_color=(0.1, 0.45, 0.8, 1),
        )
        btn_open.bind(on_release=lambda inst: _on_pdf_action(inst, 'view'))

        btn_share = MDRaisedButton(
            text='PARTAGER',
            md_bg_color=(0, 0.65, 0.25, 1),
        )
        btn_share.bind(on_release=lambda inst: _on_pdf_action(inst, 'share'))

        self.pdf_ready_dialog = MDDialog(
            title='PDF Enregistré',
            text=f'Fichier enregistré dans Téléchargements :\n{fname}',
            buttons=[btn_close, btn_open, btn_share],
        )
        self.pdf_ready_dialog.open()

    def perform_export(self):
        if getattr(self, 'settings_menu_dialog', None):
            self.settings_menu_dialog.dismiss()
        filename = f"MagPro_Stock_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
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
        except Exception as exc:
            self.notify(f'Erreur Export : {exc}', 'error')
        finally:
            conn.close()
            if os.path.exists(temp_xlsx):
                try:
                    os.remove(temp_xlsx)
                except Exception:
                    pass

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
        btns.add_widget(MDFlatButton(text='ANNULER', on_release=lambda x: self.import_diag.dismiss()))
        btns.add_widget(MDRaisedButton(text='SUIVANT', md_bg_color=(0.1, 0.4, 0.8, 1), on_release=lambda x: self._pre_process_import(self.import_file_chooser.selection)))
        content.add_widget(btns)
        self.import_diag = MDDialog(title='Importer Produits (.xlsx)', type='custom', content_cls=content, size_hint=(0.95, 0.9))
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
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            rows = list(wb.active.iter_rows(values_only=True))
            if len(rows) < 2:
                self.notify('Fichier Excel vide.', 'error')
                return
            headers = [str(h or '').strip() for h in rows[0]]
            data = rows[1:]
            if getattr(self, 'import_diag', None):
                self.import_diag.dismiss()
            self._open_mapping_dialog(headers, data)
        except Exception as exc:
            self.notify(f'Erreur lecture : {exc}', 'error')

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
        btns.add_widget(MDFlatButton(text='ANNULER', on_release=lambda x: self.mapping_diag.dismiss()))
        btns.add_widget(MDRaisedButton(text='IMPORTER', md_bg_color=(0, 0.6, 0.2, 1), on_release=lambda x: self._finalize_import(excel_headers, excel_data)))
        content.add_widget(btns)
        self.mapping_diag = MDDialog(title='Mappage Excel', type='custom', content_cls=content, size_hint=(0.95, 0.92))
        self.mapping_diag.open()

    def _finalize_import(self, excel_headers: list, excel_data: list):
        self.mapping_diag.dismiss()
        indices = {}
        for field, sp in self.spinners.items():
            if sp.text != '-- Ignorer --':
                real = self.spinner_map.get(sp.text)
                if real in excel_headers:
                    indices[field] = excel_headers.index(real)
        if 'name' not in indices:
            self.notify("Le champ 'Désignation' est obligatoire.", 'error')
            return

        added, skipped = 0, 0
        for row in excel_data:
            def _val(f, default=''):
                idx = indices.get(f)
                if idx is None or idx >= len(row) or row[idx] is None:
                    return default
                return row[idx]

            name_v = str(_val('name', '')).strip()
            if not name_v:
                continue
            try:
                self.db.save_product(
                    {
                        'action': 'add',
                        'name': name_v,
                        'barcode': str(_val('barcode', '')).strip(),
                        'product_ref': str(_val('product_ref', '')).strip() or self.db.get_next_product_ref(),
                        'reference': str(_val('reference', '')).strip(),
                        'price': FinancialEngine.to_decimal(_val('price', 0)),
                        'price_semi': FinancialEngine.to_decimal(_val('price_semi', 0)),
                        'price_wholesale': FinancialEngine.to_decimal(_val('price_wholesale', 0)),
                        'purchase_price': FinancialEngine.to_decimal(_val('purchase_price', 0)),
                        'stock_store': FinancialEngine.to_decimal(_val('stock_store', 0)),
                        'is_unlimited': False,
                    },
                    user_name=self.current_user_name,
                )
                added += 1
            except Exception:
                skipped += 1
        self.notify(f'Import terminé : {added} ajoutés ({skipped} ignorés).', 'success')
        self.check_and_load_stats()

    # --- Sauvegarde & Restauration ZIP (DB + Images) ---
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

    def perform_local_backup(self, backup_type: str = 'manual'):
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
            temp_db = os.path.join(self.user_data_dir, 'temp_backup.db')
            if os.path.exists(temp_db):
                os.remove(temp_db)

            src_conn = sqlite3.connect(self.db.db_path)
            dst_conn = sqlite3.connect(temp_db)
            src_conn.backup(dst_conn)
            dst_conn.close()
            src_conn.close()

            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                zf.write(temp_db, arcname=AppConstants.DB_NAME)
                img_dir = os.path.join(self.user_data_dir, 'product_images')
                if os.path.exists(img_dir):
                    for root, _, files in os.walk(img_dir):
                        for f in files:
                            zf.write(os.path.join(root, f), arcname=os.path.join('product_images', f))
            if os.path.exists(temp_db):
                os.remove(temp_db)
            self._rotate_backups(bdir)
            if backup_type == 'manual':
                self.notify(f'Sauvegarde enregistrée : {fname}', 'success')
        except Exception as exc:
            if backup_type == 'manual':
                self.notify(f'Erreur sauvegarde : {exc}', 'error')

    def trigger_auto_backups(self, dt=None):
        self.perform_local_backup('daily')
        self.perform_local_backup('rolling')

    def share_database_file(self):
        try:
            fname = f"MagPro_Share_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
            zip_path = self.get_unified_path(fname)
            temp_db = os.path.join(self.user_data_dir, 'temp_share.db')
            src_conn = sqlite3.connect(self.db.db_path)
            dst_conn = sqlite3.connect(temp_db)
            src_conn.backup(dst_conn)
            dst_conn.close()
            src_conn.close()

            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                zf.write(temp_db, arcname=AppConstants.DB_NAME)
                img_dir = os.path.join(self.user_data_dir, 'product_images')
                if os.path.exists(img_dir):
                    for root, _, files in os.walk(img_dir):
                        for f in files:
                            zf.write(os.path.join(root, f), arcname=os.path.join('product_images', f))
            if os.path.exists(temp_db):
                os.remove(temp_db)

            if platform == 'android':
                from jnius import autoclass, cast
                intent_cls = autoclass('android.content.Intent')
                uri_cls = autoclass('android.net.Uri')
                file_cls = autoclass('java.io.File')
                str_cls = autoclass('java.lang.String')
                strict = autoclass('android.os.StrictMode')
                strict.setVmPolicy(autoclass('android.os.StrictMode$VmPolicy$Builder')().build())

                share_i = intent_cls(intent_cls.ACTION_SEND)
                share_i.setType('application/zip')
                share_i.putExtra(intent_cls.EXTRA_STREAM, cast('android.os.Parcelable', uri_cls.fromFile(file_cls(zip_path))))
                share_i.addFlags(intent_cls.FLAG_GRANT_READ_URI_PERMISSION)
                cast('android.app.Activity', PythonActivity.mActivity).startActivity(intent_cls.createChooser(share_i, str_cls(f'Partager : {fname}')))
            else:
                self.notify(f'Archive créée : {zip_path}', 'success')
        except Exception as exc:
            self.notify(f'Erreur partage : {exc}', 'error')

    def show_restore_dialog(self):
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

        default_path = self.get_unified_path('')
        content = MDBoxLayout(orientation='vertical', spacing=10, size_hint_y=None, height=dp(520))
        bg = MDCard(orientation='vertical', size_hint_y=1, md_bg_color=(0.15, 0.15, 0.15, 1), radius=[10], padding='5dp')
        self.file_chooser = FileChooserListView(path=default_path, filters=['*.zip', '*.db'], size_hint_y=1)
        bg.add_widget(self.file_chooser)
        content.add_widget(bg)
        btns = MDBoxLayout(spacing=10, size_hint_y=None, height=dp(50))
        btns.add_widget(MDFlatButton(text='ANNULER', on_release=lambda x: self.restore_dialog.dismiss()))
        btns.add_widget(MDRaisedButton(text='RESTAURER', md_bg_color=(0.8, 0, 0, 1), on_release=lambda x: self._confirm_restore(self.file_chooser.selection[0] if self.file_chooser.selection else None)))
        content.add_widget(btns)
        self.restore_dialog = MDDialog(title='Restaurer une sauvegarde', type='custom', content_cls=content, size_hint=(0.95, 0.9))
        self.restore_dialog.open()

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

    def _confirm_restore(self, final_path: str | None):
        if not final_path or not os.path.exists(final_path):
            self.notify('Veuillez sélectionner un fichier valide.', 'error')
            return
        if getattr(self, 'restore_dialog', None):
            self.restore_dialog.dismiss()

        def _do_restore(x):
            if getattr(self, 'final_restore_confirm', None):
                self.final_restore_confirm.dismiss()
            try:
                self.db.cleanup_wal_files()
                target_db = self.db.db_path
                if final_path.lower().endswith('.zip'):
                    with zipfile.ZipFile(final_path, 'r') as zf:
                        db_members = [n for n in zf.namelist() if n.lower().endswith('.db') and '..' not in n.split('/')]
                        if not db_members:
                            raise ValueError('Archive invalide (aucune base .db trouvée).')
                        with zf.open(db_members[0]) as src, open(target_db, 'wb') as dst:
                            shutil.copyfileobj(src, dst)
                        for member in zf.namelist():
                            if member.startswith('product_images/') and not os.path.isabs(member) and '..' not in member.split('/'):
                                zf.extract(member, self.user_data_dir)
                else:
                    shutil.copyfile(final_path, target_db)
                self.db = DatabaseManager()
                self.is_seller_mode = self.db.get_bool_setting('config_seller_mode', default=False)
                self.update_dashboard_layout()
                self.notify('Restauration réussie !', 'success')
            except Exception as exc:
                self.notify(f'Erreur restauration : {exc}', 'error')

        self.final_restore_confirm = MDDialog(
            title='Attention',
            text='Voulez-vous vraiment restaurer cette sauvegarde ?\nToutes les données actuelles seront remplacées.',
            buttons=[
                MDFlatButton(text='ANNULER', on_release=lambda x: self.final_restore_confirm.dismiss()),
                MDRaisedButton(text='OUI, RESTAURER', md_bg_color=(0.8, 0, 0, 1), on_release=_do_restore),
            ],
        )
        self.final_restore_confirm.open()

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
