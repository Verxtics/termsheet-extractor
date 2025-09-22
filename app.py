import streamlit as st
import tempfile
import os
import pandas as pd
import pdfplumber
from datetime import datetime
import re
import warnings
import shutil
warnings.filterwarnings('ignore')

class FixedIncomeTermsheetExtractor:
    def __init__(self):
        # Define your exact column mapping from the Database sheet
        self.column_mapping = {
            0: 'Investment Name',
            1: 'Issuer', 
            2: 'Product Name',
            3: 'Investment Thematic',
            4: 'TYPE',
            5: 'Coupon - QTR',
            6: 'Coupon Rate - Annual',
            7: 'Product Status',
            8: 'Knock-In%',
            9: 'Knock-Out%',
            10: 'Minimum Tenor (Q)',
            11: 'Maximum Tenor (Q)',
            12: 'Observation Frequency',
            13: 'CCY',
            14: 'Strike Date',
            15: 'Issue Date',
            16: 'ISIN',
            17: None,  # Empty column
            18: 'Underlying 1',
            19: 'Underlying 2',
            20: 'Underlying 3',
            21: 'Underlying 4',
            22: None,  # Empty column
            23: 'Investment $',
            24: 'Total Units ',
            25: 'Notional Value',
            26: 'AUD Equivalent',
            27: 'Maturity Date',
            28: 'Revenue',
            29: None,  # Empty column
            30: 'Underlying 1 - Issue Price',
            31: 'Underlying 2 - Issue Price',
            32: 'Underlying 3 - Issue Price',
            33: 'Underlying 4 - Issue Price',
            34: None,  # Empty column
            35: 'Underlying 1 - Knock In Price',
            36: 'Underlying 2 - Knock In Price',
            37: 'Underlying 3 - Knock In Price',
            38: 'Underlying 4 - Knock In Price',
            39: None,  # Empty column
            40: 'Underlying 1 - Knock Out',
            41: 'Underlying 2 - Knock Out',
            42: 'Underlying 3 - Knock Out',
            43: 'Underlying 4 - Knock Out'
        }
        
        # Define issuer-specific patterns
        self.issuer_patterns = {
            'morgan_stanley': {
                'identifiers': ['MORGAN STANLEY', 'MS&Co', 'Morgan Stanley & Co'],
                'issuer_name': 'Morgan Stanley & Co. International PLC',
                'patterns': {
                    'isin': r'[A-Z]{2}[A-Z0-9]{10}',
                    'coupon_rate': r'(\d+\.?\d*)\s*%.*(?:coupon|rate)',
                    'knock_in': r'(\d+)%.*(?:knock.?in|barrier)',
                    'dates': r'(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})'
                }
            },
            'macquarie': {
                'identifiers': ['MACQUARIE', 'MBL', 'Macquarie Bank Limited', 'EQUITY LINKED NOTE'],
                'issuer_name': 'Macquarie Bank Limited',
                'patterns': {
                    'isin': r'[A-Z]{2}[A-Z0-9]{10}',
                    'product_name': r'(EQUITY LINKED NOTE)',
                    'coupon_base_rate': r'(\d+\.?\d*%)',  # Base rate for escalation
                    'coupon_formula': r'(\d+\.?\d*%)\s*x\s*\(1\s*\+\s*Number of Periods\)',
                    'knock_in_price': r'Knock-in Price.*?(\d+\.?\d*)%.*?Initial Price',
                    'knock_out_price': r'Knock-out Price.*?(\d+\.?\d*)%.*?Initial Price',
                    'tickers': r'([A-Z]{3,4}\.[A-Z]{1,2})',  # Exchange-specific tickers
                    'usd_prices': r'\[USD\s*([\d.]+)\]',
                    'aud_amounts': r'AUD\s*([\d,]+(?:\.\d{2})?)',
                    'dates': r'(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})',
                    'aggregate_nominal': r'Aggregate Nominal Amount.*?AUD\s*([\d,]+(?:\.\d{2})?)',
                    'denomination': r'Specified Denomination.*?AUD\s*([\d,]+)',
                    'us_tech_stocks': r'(ORCL\.N|AVGO\.OQ|META\.OQ|NVDA\.OQ)'
                }
            },
            'citigroup': {
                'identifiers': ['CITIGROUP', 'CITI', 'Citigroup Global Markets Holdings', 'CGMHI', 'Snowballing Autocall Notes'],
                'issuer_name': 'Citigroup Global Markets Holdings Inc.',
                'patterns': {
                    'isin': r'[A-Z]{2}[A-Z0-9]{10}',
                    'product_name': r'(Snowballing Autocall Notes[^"]*)',
                    'snowball_percentage': r'Snowball Percentage.*?(\d+\.?\d*%)',
                    'coupon_rates': r'(\d+\.?\d*%(?:,\s*\d+\.?\d*%)*)',  # Multiple escalating rates
                    'knock_in_barrier': r'Knock-In Barrier Level.*?(\d+\.?\d*)%',
                    'autocall_barrier': r'Autocall Barrier Level.*?(\d+\.?\d*)%',
                    'initial_level': r'Initial Level.*?([A-Z]{3}\s*[\d,]+\.?\d*)',
                    'currency': r'Currency.*?(Australian Dollar|USD|EUR|GBP|CHF|AUD)',
                    'denomination': r'Denomination.*?([A-Z]{3}\s*[\d,]+)',
                    'issue_size': r'Issue Size.*?([A-Z]{3}\s*[\d,]+)',
                    'dates': r'(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})',
                    'underlying_names': r'(Banco Santander SA|BNP Paribas|Societe Generale|UBS Group AG)',
                    'worst_performing': r'Worst Performing.*?Underlying'
                }
            },
            'goldman_sachs': {
                'identifiers': ['GOLDMAN SACHS', 'GS&Co', 'Goldman Sachs'],
                'issuer_name': 'Goldman Sachs International',
                'patterns': {
                    'isin': r'[A-Z]{2}[A-Z0-9]{10}',
                    'coupon_rate': r'(\d+\.?\d*)\s*%.*(?:coupon|quarterly)',
                    'knock_in': r'(\d+)%.*(?:protection|barrier)',
                    'dates': r'(\d{1,2})/(\d{1,2})/(\d{4})'
                }
            },
            'ubs': {
                'identifiers': ['UBS', 'UBS Investments Australia', 'UBS AG', 'Callable Equity Basket', 'UBS Equity Goals'],
                'issuer_name': 'UBS Investments Australia Pty Ltd',
                'patterns': {
                    'isin': r'[A-Z]{2}[A-Z0-9]{10}',
                    'product_name': r'(Callable Equity Basket[^"]*|UBS Equity Goals)',
                    'kick_in_level': r'Kick-in Level.*?(\d+)%.*?Initial Level',
                    'call_level': r'Call Level.*?(\d+)%.*?Initial Level',
                    'snowball_coupon_rate': r'Snowball Coupon Rate.*?(\d+\.?\d*)%',
                    'initial_level': r'Initial Level.*?USD\s*([\d.]+)',
                    'usd_4decimal': r'USD\s*([\d.]{4,})',  # 4 decimal place amounts
                    'bloomberg_codes': r'Bloomberg code:\s*([A-Z]+\s+[A-Z]{2})',
                    'dates': r'(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})',
                    'aud_amounts': r'AUD\s*([\d,]+)',
                    'tech_stocks': r'(Alphabet Inc|Meta Platforms Inc|Microsoft Corporation|Oracle Corporation)',
                    'call_observation_n': r'N.*?(\d+)'  # For tenor extraction
                }
            },
            'bnp_paribas': {
                'identifiers': ['BNP PARIBAS', 'BNP Paribas Issuance', 'Stock Basket Periodic Callable', 'Certificate'],
                'issuer_name': 'BNP Paribas Issuance B.V.',
                'patterns': {
                    'isin': r'[A-Z]{2}[A-Z0-9]{10}',
                    'product_title': r'(\d+\s+Months.*?Certificates)',
                    'coupon_annual': r'C\s*=\s*(\d+)%\s*p\.a\.',
                    'knock_in_percentage': r'(\d+\.?\d*)%.*?Initial Spot Price',
                    'trigger_percentage': r'(\d+)%.*?Initial Spot Price',
                    'initial_spot_price': r'Initial Spot Price.*?USD\s*([\d.]+)',
                    'knock_in_price': r'Knock-In Price.*?USD\s*([\d.]+)',
                    'trigger_price': r'Trigger Price.*?USD\s*([\d.]+)',
                    'bloomberg_tickers': r'([A-Z]+\s+UW)',
                    'company_names': r'(Alphabet Inc|Meta Platforns Inc|NVIDIA Corp|MICROSOFT CORP)',
                    'dates_ordinal': r'(\w+)\s+(\d{1,2})(?:st|nd|rd|th),\s+(\d{4})',
                    'aud_amounts': r'AUD\s*([\d,]+)',
                    'certificate_count': r'Number of Certificates.*?(\d+)',
                    'notional_per_cert': r'AUD\s*(\d+).*?per Certificate',
                    'reference_code': r'(CE\d+[A-Z]+)'
                }
            },
            'barclays': {
                'identifiers': ['BARCLAYS', 'Barclays Bank PLC', 'Periodic Snowball Autocall', 'Quanto AUD'],
                'issuer_name': 'Barclays Bank PLC',
                'patterns': {
                    'isin': r'[A-Z]{2}[A-Z0-9]{10}',
                    'product_name': r'(Periodic Snowball Autocall|Quanto AUD[^"]*)',
                    'coupon_quarterly': r'(\d+\.?\d*)%.*?per quarter',
                    'final_coupon': r'(\d+\.?\d*)%.*?final',
                    'knock_in_event': r'Knock-in Event.*?(\d+)%',
                    'autocall_trigger': r'Autocall Trigger.*?(\d+)%',
                    'initial_price': r'Initial Price.*?USD\s*([\d.]+)',
                    'specified_denomination': r'Specified Denomination.*?AUD\s*([\d,]+)',
                    'aggregate_nominal': r'Aggregate Nominal Amount.*?AUD\s*([\d,]+)',
                    'dates': r'(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})',
                    'initial_valuation_date': r'Initial Valuation Date.*?(\d{1,2}\s+\w+\s+\d{4})',
                    'us_tech_stocks': r'(ALPHABET INC|MICROSOFT CORP|META PLATFORMS|NVIDIA CORP)',
                    'bloomberg_usd': r'([A-Z]+\s+UW)',
                    'final_settlement_amount': r'Final Cash Settlement Amount'
                }
            },
            'natixis': {
                'identifiers': ['NATIXIS', 'EMTN', 'Autocall Incremental', 'no-Knock-In-Coupon'],
                'issuer_name': 'NATIXIS',
                'patterns': {
                    'isin': r'[A-Z]{2}[A-Z0-9]{10}',
                    'product_name': r'(EMTN.*?Autocall Incremental|Autocall Incremental[^"]*)',
                    'coupon_quarterly': r'(\d+\.?\d*)%.*?quarterly',
                    'automatic_early_redemption': r'Automatic Early Redemption Rate.*?(\d+\.?\d*)%',
                    'knock_in_event': r'Knock-in Event.*?(\d+)%',
                    'autocall_percentage': r'(\d+)%.*?Initial Price',
                    'initial_price': r'Initial Price.*?(EUR|GBp|CHF)\s*([\d.]+)',
                    'denomination': r'Denomination.*?AUD\s*([\d,]+)',
                    'aggregate_nominal_lower': r'Aggregate nominal amount.*?AUD\s*([\d,]+)',
                    'dates': r'(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})',
                    'strike_date': r'Strike Date.*?(\d{1,2}\s+\w+\s+\d{4})',
                    'european_banks': r'(Banco Bilbao|Barclays PLC|UBS Group|Societe Generale)',
                    'mixed_currency': r'(EUR|GBp|CHF)\s*[\d.]+',
                    'final_redemption_amount': r'Final Redemption Amount',
                    'lowest_performing': r'Lowest Performing Share'
                }
            },
            'generic': {
                'identifiers': [],
                'issuer_name': 'Unknown Issuer',
                'patterns': {
                    'isin': r'[A-Z]{2}[A-Z0-9]{10}',
                    'coupon_rate': r'(\d+\.?\d*)\s*%',
                    'knock_in': r'(\d+)%',
                    'dates': r'(\d{1,2})[\/\-\s](\d{1,2})[\/\-\s](\d{4})'
                }
            }
        }

    def detect_issuer(self, text):
        """Detect which issuer based on text content"""
        text_upper = text.upper()
        
        for issuer_key, config in self.issuer_patterns.items():
            if issuer_key == 'generic':
                continue
            for identifier in config['identifiers']:
                if identifier.upper() in text_upper:
                    return issuer_key
        
        return 'generic'

    def extract_with_patterns(self, text, patterns, issuer_type):
        """Extract data using regex patterns with issuer-specific logic"""
        extracted = {}
        
        # Extract ISIN
        isin_match = re.search(patterns['isin'], text)
        extracted['ISIN'] = isin_match.group(0) if isin_match else ''
        
        # Handle issuer-specific coupon rate extraction
        if issuer_type == 'citigroup':
            # Extract escalating coupon rates for Citi's snowball structure
            coupon_rates_text = re.findall(r'(\d+\.?\d*%)', text)
            if coupon_rates_text:
                rates = [float(rate.replace('%', '')) for rate in coupon_rates_text if float(rate.replace('%', '')) > 1]
                if rates:
                    max_rate = max(rates) / 100  # Convert to decimal
                    extracted['Coupon Rate - Annual'] = max_rate
                else:
                    extracted['Coupon Rate - Annual'] = ''
            else:
                extracted['Coupon Rate - Annual'] = ''
                
        elif issuer_type == 'macquarie':
            # Extract base coupon rate for Macquarie's escalation formula
            coupon_match = re.search(patterns.get('coupon_base_rate', ''), text, re.IGNORECASE)
            if coupon_match:
                rate_value = float(coupon_match.group(1).replace('%', '')) / 100
                extracted['Coupon Rate - Annual'] = rate_value
            else:
                extracted['Coupon Rate - Annual'] = ''
                
        elif issuer_type == 'ubs':
            # Extract UBS snowball coupon rate
            snowball_match = re.search(patterns.get('snowball_coupon_rate', ''), text, re.IGNORECASE)
            if snowball_match:
                rate_value = float(snowball_match.group(1)) / 100
                extracted['Coupon Rate - Annual'] = rate_value
            else:
                extracted['Coupon Rate - Annual'] = ''
                
        elif issuer_type == 'bnp_paribas':
            # Extract BNP annual coupon rate from formula
            coupon_match = re.search(patterns.get('coupon_annual', ''), text, re.IGNORECASE)
            if coupon_match:
                rate_value = float(coupon_match.group(1)) / 100
                extracted['Coupon Rate - Annual'] = rate_value
            else:
                extracted['Coupon Rate - Annual'] = ''
                
        elif issuer_type == 'barclays':
            # Extract Barclays quarterly coupon rate
            quarterly_match = re.search(patterns.get('coupon_quarterly', ''), text, re.IGNORECASE)
            if quarterly_match:
                quarterly_rate = float(quarterly_match.group(1)) / 100
                # Convert to annual (quarterly * 4)
                extracted['Coupon Rate - Annual'] = quarterly_rate * 4
            else:
                # Try final coupon pattern
                final_match = re.search(patterns.get('final_coupon', ''), text, re.IGNORECASE)
                if final_match:
                    final_rate = float(final_match.group(1)) / 100
                    extracted['Coupon Rate - Annual'] = final_rate
                else:
                    extracted['Coupon Rate - Annual'] = ''
                    
        elif issuer_type == 'natixis':
            # Extract Natixis quarterly coupon rate
            quarterly_match = re.search(patterns.get('coupon_quarterly', ''), text, re.IGNORECASE)
            if quarterly_match:
                quarterly_rate = float(quarterly_match.group(1)) / 100
                # Convert to annual (quarterly * 4)
                extracted['Coupon Rate - Annual'] = quarterly_rate * 4
            else:
                # Try automatic early redemption rate
                auto_match = re.search(patterns.get('automatic_early_redemption', ''), text, re.IGNORECASE)
                if auto_match:
                    rate_value = float(auto_match.group(1)) / 100
                    extracted['Coupon Rate - Annual'] = rate_value
                else:
                    extracted['Coupon Rate - Annual'] = ''
        else:
            # Standard coupon rate extraction for other issuers
            coupon_match = re.search(patterns.get('coupon_rate', ''), text, re.IGNORECASE)
            if coupon_match:
                rate_value = float(coupon_match.group(1)) / 100  # Convert to decimal
                extracted['Coupon Rate - Annual'] = rate_value
            else:
                extracted['Coupon Rate - Annual'] = ''
        
        # Handle issuer-specific knock-in barrier extraction
        if issuer_type == 'citigroup':
            knock_in_match = re.search(patterns.get('knock_in_barrier', ''), text, re.IGNORECASE)
            if knock_in_match:
                barrier_value = float(knock_in_match.group(1)) / 100
                extracted['Knock-In%'] = barrier_value
            else:
                extracted['Knock-In%'] = 0.6  # Default 60% for Citi
                
        elif issuer_type == 'macquarie':
            knock_in_match = re.search(patterns.get('knock_in_price', ''), text, re.IGNORECASE)
            if knock_in_match:
                barrier_value = float(knock_in_match.group(1)) / 100
                extracted['Knock-In%'] = barrier_value
            else:
                extracted['Knock-In%'] = 0.6  # Default 60% for MBL
                
        elif issuer_type == 'ubs':
            kick_in_match = re.search(patterns.get('kick_in_level', ''), text, re.IGNORECASE)
            if kick_in_match:
                barrier_value = float(kick_in_match.group(1)) / 100
                extracted['Knock-In%'] = barrier_value
            else:
                extracted['Knock-In%'] = 0.6  # Default 60% for UBS
                
        elif issuer_type == 'bnp_paribas':
            knock_in_match = re.search(patterns.get('knock_in_percentage', ''), text, re.IGNORECASE)
            if knock_in_match:
                barrier_value = float(knock_in_match.group(1)) / 100
                extracted['Knock-In%'] = barrier_value
            else:
                extracted['Knock-In%'] = 0.6  # Default 60% for BNP
                
        elif issuer_type == 'barclays':
            knock_in_match = re.search(patterns.get('knock_in_event', ''), text, re.IGNORECASE)
            if knock_in_match:
                barrier_value = float(knock_in_match.group(1)) / 100
                extracted['Knock-In%'] = barrier_value
            else:
                extracted['Knock-In%'] = 0.6  # Default 60% for Barclays
                
        elif issuer_type == 'natixis':
            knock_in_match = re.search(patterns.get('knock_in_event', ''), text, re.IGNORECASE)
            if knock_in_match:
                barrier_value = float(knock_in_match.group(1)) / 100
                extracted['Knock-In%'] = barrier_value
            else:
                extracted['Knock-In%'] = 0.6  # Default 60% for Natixis
        else:
            # Standard knock-in barrier extraction
            knock_in_match = re.search(patterns.get('knock_in', ''), text, re.IGNORECASE)
            if knock_in_match:
                barrier_value = float(knock_in_match.group(1)) / 100  # Convert to decimal
                extracted['Knock-In%'] = barrier_value
            else:
                extracted['Knock-In%'] = ''
        
        # Handle issuer-specific knock-out/autocall extraction
        if issuer_type == 'citigroup':
            autocall_match = re.search(patterns.get('autocall_barrier', ''), text, re.IGNORECASE)
            if autocall_match:
                autocall_value = float(autocall_match.group(1)) / 100
                extracted['Knock-Out%'] = autocall_value
            else:
                extracted['Knock-Out%'] = 0.9  # Default 90% for Citi
                
        elif issuer_type == 'macquarie':
            knock_out_match = re.search(patterns.get('knock_out_price', ''), text, re.IGNORECASE)
            if knock_out_match:
                autocall_value = float(knock_out_match.group(1)) / 100
                extracted['Knock-Out%'] = autocall_value
            else:
                extracted['Knock-Out%'] = 0.9  # Default 90% for MBL
                
        elif issuer_type == 'ubs':
            call_match = re.search(patterns.get('call_level', ''), text, re.IGNORECASE)
            if call_match:
                autocall_value = float(call_match.group(1)) / 100
                extracted['Knock-Out%'] = autocall_value
            else:
                extracted['Knock-Out%'] = 0.9  # Default 90% for UBS
                
        elif issuer_type == 'bnp_paribas':
            trigger_match = re.search(patterns.get('trigger_percentage', ''), text, re.IGNORECASE)
            if trigger_match:
                autocall_value = float(trigger_match.group(1)) / 100
                extracted['Knock-Out%'] = autocall_value
            else:
                extracted['Knock-Out%'] = 0.9  # Default 90% for BNP
                
        elif issuer_type == 'barclays':
            autocall_match = re.search(patterns.get('autocall_trigger', ''), text, re.IGNORECASE)
            if autocall_match:
                autocall_value = float(autocall_match.group(1)) / 100
                extracted['Knock-Out%'] = autocall_value
            else:
                extracted['Knock-Out%'] = 0.9  # Default 90% for Barclays
                
        elif issuer_type == 'natixis':
            autocall_match = re.search(patterns.get('autocall_percentage', ''), text, re.IGNORECASE)
            if autocall_match:
                autocall_value = float(autocall_match.group(1)) / 100
                extracted['Knock-Out%'] = autocall_value
            else:
                extracted['Knock-Out%'] = 0.9  # Default 90% for Natixis
        else:
            # Standard knock-out extraction (if pattern exists)
            extracted['Knock-Out%'] = 0.95  # Default assumption
        
        # Extract dates using issuer-specific formats
        if issuer_type == 'bnp_paribas':
            # Handle BNP's ordinal date format (February 3rd, 2025)
            dates = re.findall(patterns.get('dates_ordinal', ''), text)
            extracted['dates_found'] = dates
        else:
            # Standard date extraction
            dates = re.findall(patterns.get('dates', patterns['dates']), text)
            extracted['dates_found'] = dates
        
        return extracted

    def extract_currency(self, text, issuer_type):
        """Extract currency from text with issuer-specific handling"""
        
        # Citigroup-specific currency extraction
        if issuer_type == 'citigroup':
            # Look for "Australian Dollar (AUD)" pattern
            aud_match = re.search(r'Australian Dollar.*?\(AUD\)', text, re.IGNORECASE)
            if aud_match:
                return 'AUD'
            
            # Look for "Currency" field
            currency_match = re.search(r'Currency.*?(AUD|USD|EUR|GBP|CHF)', text, re.IGNORECASE)
            if currency_match:
                return currency_match.group(1).upper()
        
        # Standard currency patterns
        currency_patterns = [
            r'\b(USD|EUR|GBP|JPY|AUD|CAD|CHF|HKD|SGD)\b',
            r'\$([\d,]+)',  # Dollar amounts
            r'€([\d,]+)',   # Euro amounts
            r'£([\d,]+)'    # Pound amounts
        ]
        
        for pattern in currency_patterns:
            match = re.search(pattern, text)
            if match:
                if pattern.startswith(r'\b'):
                    return match.group(1)
                elif pattern.startswith(r'\$'):
                    return 'USD'
                elif pattern.startswith(r'€'):
                    return 'EUR'
                elif pattern.startswith(r'£'):
                    return 'GBP'
        
        return 'USD'  # Default assumption

    def extract_notional_amount(self, text, issuer_type):
        """Extract notional amount with issuer-specific handling"""
        
        if issuer_type == 'citigroup':
            # Look for "Issue Size" pattern
            issue_size_match = re.search(r'Issue Size.*?AUD\s*([\d,]+)', text, re.IGNORECASE)
            if issue_size_match:
                return int(issue_size_match.group(1).replace(',', ''))
            
            # Look for "Denomination" pattern  
            denomination_match = re.search(r'Denomination.*?AUD\s*([\d,]+)', text, re.IGNORECASE)
            if denomination_match:
                return int(denomination_match.group(1).replace(',', ''))
                
        elif issuer_type == 'macquarie':
            # Look for "Aggregate Nominal Amount"
            aggregate_match = re.search(r'Aggregate Nominal Amount.*?AUD\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
            if aggregate_match:
                return int(float(aggregate_match.group(1).replace(',', '')))
                
        elif issuer_type == 'ubs':
            # Look for "Issue proceeds" or similar
            proceeds_match = re.search(r'(?:Issue proceeds|Issue Amount).*?AUD\s*([\d,]+)', text, re.IGNORECASE)
            if proceeds_match:
                return int(proceeds_match.group(1).replace(',', ''))
                
        elif issuer_type == 'bnp_paribas':
            # Look for "Issue Amount"
            issue_amount_match = re.search(r'Issue Amount.*?AUD\s*([\d,]+)', text, re.IGNORECASE)
            if issue_amount_match:
                return int(issue_amount_match.group(1).replace(',', ''))
                
        elif issuer_type == 'barclays':
            # Look for "Aggregate Nominal Amount"
            aggregate_match = re.search(r'Aggregate Nominal Amount.*?AUD\s*([\d,]+)', text, re.IGNORECASE)
            if aggregate_match:
                return int(aggregate_match.group(1).replace(',', ''))
            
            # Look for "Specified Denomination"
            denomination_match = re.search(r'Specified Denomination.*?AUD\s*([\d,]+)', text, re.IGNORECASE)
            if denomination_match:
                return int(denomination_match.group(1).replace(',', ''))
                
        elif issuer_type == 'natixis':
            # Look for "Aggregate nominal amount" (lowercase)
            aggregate_match = re.search(r'Aggregate nominal amount.*?AUD\s*([\d,]+)', text, re.IGNORECASE)
            if aggregate_match:
                return int(aggregate_match.group(1).replace(',', ''))
            
            # Look for "Denomination"
            denomination_match = re.search(r'Denomination.*?AUD\s*([\d,]+)', text, re.IGNORECASE)
            if denomination_match:
                return int(denomination_match.group(1).replace(',', ''))
        
        # Standard notional patterns for other issuers
        notional_patterns = [
            r'notional.*?([0-9,]+)',
            r'principal.*?([0-9,]+)',
            r'amount.*?([0-9,]+)',
            r'issue size.*?([0-9,]+)'
        ]
        
        for pattern in notional_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return int(match.group(1).replace(',', ''))
        
        return ''

    def extract_underlying_assets(self, text, tables_data, issuer_type):
        """Extract underlying asset information with issuer-specific handling"""
        underlyings = []
        
        # Issuer-specific underlying extraction
        if issuer_type == 'citigroup':
            # Citigroup European banks
            citi_banks = ['Banco Santander SA', 'BNP Paribas', 'Societe Generale', 'UBS Group AG']
            for bank in citi_banks:
                if bank.upper() in text.upper():
                    underlyings.append({
                        'Name': bank,
                        'Initial_Price': '',
                        'Knock_In_Price': '',
                        'Knock_Out_Price': ''
                    })
                    
        elif issuer_type == 'macquarie':
            # Macquarie US Tech stocks with tickers
            mbl_tech_stocks = [
                ('Oracle Corporation', 'ORCL.N'),
                ('Broadcom Inc.', 'AVGO.OQ'),
                ('Meta Platforms', 'META.OQ'),
                ('NVIDIA Corporation', 'NVDA.OQ')
            ]
            
            for company, ticker in mbl_tech_stocks:
                if ticker in text or company.upper() in text.upper():
                    underlyings.append({
                        'Name': company,
                        'Ticker': ticker,
                        'Initial_Price': '',
                        'Knock_In_Price': '',
                        'Knock_Out_Price': ''
                    })
                    
        elif issuer_type == 'ubs':
            # UBS US Tech stocks with Bloomberg codes
            ubs_tech_stocks = [
                ('Alphabet Inc', 'GOOG UW'),
                ('Meta Platforms Inc', 'META UW'),
                ('Microsoft Corporation', 'MSFT UW'),
                ('Oracle Corporation', 'ORCL UN')
            ]
            
            for company, bloomberg in ubs_tech_stocks:
                if company in text or bloomberg in text:
                    underlyings.append({
                        'Name': company,
                        'Bloomberg_Code': bloomberg,
                        'Initial_Price': '',
                        'Knock_In_Price': '',
                        'Knock_Out_Price': ''
                    })
                    
        elif issuer_type == 'bnp_paribas':
            # BNP US Tech stocks with Bloomberg tickers
            bnp_tech_stocks = [
                ('Alphabet Inc', 'GOOGL UW'),
                ('Meta Platforns Inc', 'META UW'),  # Note: typo in original
                ('NVIDIA Corp', 'NVDA UW'),
                ('MICROSOFT CORP', 'MSFT UW')
            ]
            
            for company, bloomberg in bnp_tech_stocks:
                if company in text or bloomberg in text:
                    underlyings.append({
                        'Name': company,
                        'Bloomberg_Code': bloomberg,
                        'Initial_Price': '',
                        'Knock_In_Price': '',
                        'Knock_Out_Price': ''
                    })
                    
        elif issuer_type == 'barclays':
            # Barclays US Tech stocks
            barclays_tech_stocks = [
                ('ALPHABET INC-CL A', 'GOOGL UW'),
                ('MICROSOFT CORP', 'MSFT UW'),
                ('META PLATFORMS INC-CLASS A', 'META UW'),
                ('NVIDIA CORP', 'NVDA UW')
            ]
            
            for company, bloomberg in barclays_tech_stocks:
                if company in text or bloomberg in text:
                    underlyings.append({
                        'Name': company,
                        'Bloomberg_Code': bloomberg,
                        'Initial_Price': '',
                        'Knock_In_Price': '',
                        'Knock_Out_Price': ''
                    })
                    
        elif issuer_type == 'natixis':
            # Natixis European banks
            natixis_banks = [
                ('Banco Bilbao Vizcaya Argentaria SA', 'BBVA SQ'),
                ('Barclays PLC', 'BARC LN'),
                ('UBS Group AG', 'UBSG SE'),
                ('Societe Generale SA', 'GLE FP')
            ]
            
            for bank, bloomberg in natixis_banks:
                if bank in text or bloomberg in text:
                    underlyings.append({
                        'Name': bank,
                        'Bloomberg_Code': bloomberg,
                        'Initial_Price': '',
                        'Knock_In_Price': '',
                        'Knock_Out_Price': ''
                    })
        else:
            # Standard underlying extraction for other issuers
            stock_patterns = [
                r'([A-Z][A-Z\s&]+(?:INC|CORP|LTD|CO|GROUP))',
                r'Bloomberg.*?([A-Z0-9\s]+Equity)',
                r'Ticker.*?([A-Z]{2,6})'
            ]
            
            for pattern in stock_patterns:
                matches = re.findall(pattern, text)
                for match in matches[:4]:  # Limit to 4 underlyings
                    underlyings.append({
                        'Name': match.strip(),
                        'Initial_Price': '',
                        'Knock_In_Price': '',
                        'Knock_Out_Price': ''
                    })
        
        # Try to extract prices from tables for all issuers
        for table in tables_data:
            if len(table) > 1:
                headers = [str(cell).upper() if cell else '' for cell in table[0]]
                # Look for price-related headers
                if any(keyword in ' '.join(headers) for keyword in ['INITIAL', 'PRICE', 'LEVEL', 'UNDERLYING', 'SPOT']):
                    for i, row in enumerate(table[1:]):
                        if row and len(row) > 0 and i < len(underlyings):
                            # Extract prices based on issuer format
                            for j, cell in enumerate(row):
                                if cell and isinstance(cell, str):
                                    # USD price patterns
                                    if re.match(r'USD\s*[\d.]+', str(cell)):
                                        if 'initial' in headers[j].lower() or 'spot' in headers[j].lower():
                                            underlyings[i]['Initial_Price'] = str(cell)
                                        elif 'knock' in headers[j].lower() or 'kick' in headers[j].lower():
                                            underlyings[i]['Knock_In_Price'] = str(cell)
                                        elif 'call' in headers[j].lower() or 'trigger' in headers[j].lower():
                                            underlyings[i]['Knock_Out_Price'] = str(cell)
                                    # EUR/CHF price patterns (for Citi)
                                    elif re.match(r'(EUR|CHF)\s*[\d.]+', str(cell)):
                                        if not underlyings[i]['Initial_Price']:
                                            underlyings[i]['Initial_Price'] = str(cell)
        
        return underlyings[:4]  # Limit to 4 underlyings

    def extract_termsheet_data(self, pdf_path):
        """Main extraction function"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                full_text = ""
                tables_data = []
                
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        full_text += page_text + "\n"
                    
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) > 1:
                            tables_data.append(table)
                            
        except Exception as e:
            return {'error': f'Failed to read PDF: {str(e)}'}
        
        # Detect issuer
        issuer_type = self.detect_issuer(full_text)
        issuer_config = self.issuer_patterns[issuer_type]
        
        # Extract using detected issuer patterns
        extracted = self.extract_with_patterns(full_text, issuer_config['patterns'], issuer_type)
        
        # Core information
        extracted['Issuer'] = issuer_config['issuer_name']
        extracted['CCY'] = self.extract_currency(full_text, issuer_type)
        extracted['Notional Value'] = self.extract_notional_amount(full_text, issuer_type)
        
        # Extract underlying assets
        underlying_assets = self.extract_underlying_assets(full_text, tables_data, issuer_type)
        
        # Extract dates more intelligently
        dates_found = extracted.get('dates_found', [])
        if dates_found:
            extracted['Issue Date'] = f"{dates_found[0][0]}/{dates_found[0][1]}/{dates_found[0][2]}" if dates_found else ''
            extracted['Strike Date'] = f"{dates_found[1][0]}/{dates_found[1][1]}/{dates_found[1][2]}" if len(dates_found) > 1 else ''
            extracted['Maturity Date'] = f"{dates_found[-1][0]}/{dates_found[-1][1]}/{dates_found[-1][2]}" if dates_found else ''
        
        # Add metadata
        extracted['Source_File'] = os.path.basename(pdf_path)
        extracted['Detected_Issuer_Type'] = issuer_type
        extracted['underlying_assets'] = underlying_assets
        
        return extracted

def create_database_row(data):
    """Convert extracted data to match your exact Database sheet format"""
    
    # Create a row with 44 columns to match your Database sheet
    row = [''] * 44
    
    # Handle different issuer-specific product naming
    issuer_type = data.get('Detected_Issuer_Type', '')
    
    if issuer_type == 'citigroup':
        investment_name = 'CG ' + data.get('Maturity Date', '').replace(' ', '-') if data.get('Maturity Date') else 'CG Product'
        product_name = 'Snowballing Autocall Notes'
        investment_thematic = 'European Banking'
        product_type = 'Snowball AC'
        min_tenor = 6  # Citi starts at ~1.5 years
        
    elif issuer_type == 'macquarie':
        investment_name = 'MBL ' + data.get('Maturity Date', '').replace(' ', '-') if data.get('Maturity Date') else 'MBL Product'
        product_name = 'Equity Linked Note'
        investment_thematic = 'US Tech'
        product_type = 'ACE 90%'  # Based on 90% knock-out level
        min_tenor = 6  # MBL starts at ~1.5 years
        
    elif issuer_type == 'ubs':
        investment_name = 'UBS ' + data.get('Maturity Date', '').replace(' ', '-') if data.get('Maturity Date') else 'UBS Product'
        product_name = 'UBS Equity Goals'
        investment_thematic = 'US Tech'
        product_type = 'Snowball AC'  # UBS uses snowball coupon
        min_tenor = 2  # UBS starts from N=2
        
    elif issuer_type == 'bnp_paribas':
        investment_name = 'BNP ' + data.get('Maturity Date', '').replace(' ', '-') if data.get('Maturity Date') else 'BNP Product'
        product_name = 'Certificate'
        investment_thematic = 'US Tech'
        product_type = 'ACE 90%'  # Based on 90% trigger level
        min_tenor = 12  # BNP 36 months product
        
    elif issuer_type == 'barclays':
        investment_name = 'BARC ' + data.get('Maturity Date', '').replace(' ', '-') if data.get('Maturity Date') else 'BARC Product'
        product_name = 'Periodic Snowball Autocall'
        investment_thematic = 'US Tech'
        product_type = 'Snowball AC'  # Barclays snowball structure
        min_tenor = 4  # Standard quarterly structure
        
    elif issuer_type == 'natixis':
        investment_name = 'NATIXIS ' + data.get('Maturity Date', '').replace(' ', '-') if data.get('Maturity Date') else 'NATIXIS Product'
        product_name = 'Autocall Incremental'
        investment_thematic = 'European Banking'
        product_type = 'ACE 90%'  # Based on 90% autocall level
        min_tenor = 4  # Standard quarterly structure
        
    else:
        investment_name = data.get('Source_File', '').replace('.pdf', '')
        product_name = f"{issuer_type.replace('_', ' ').title()} Product"
        investment_thematic = 'Structured Product'
        product_type = 'Autocallable'
        min_tenor = 4
    
    # Map data to your exact column positions
    row[0] = investment_name  # Investment Name
    row[1] = data.get('Issuer', '')  # Issuer
    row[2] = product_name  # Product Name
    row[3] = investment_thematic  # Investment Thematic
    row[4] = product_type  # TYPE
    row[5] = 'Quarterly'  # Coupon - QTR
    row[6] = data.get('Coupon Rate - Annual', '')  # Coupon Rate - Annual
    row[7] = 'Active'  # Product Status
    row[8] = data.get('Knock-In%', '')  # Knock-In%
    row[9] = data.get('Knock-Out%', 0.95)  # Knock-Out%
    row[10] = min_tenor  # Minimum Tenor (Q) - issuer-specific
    row[11] = 12  # Maximum Tenor (Q)
    row[12] = 'Quarterly'  # Observation Frequency
    row[13] = data.get('CCY', 'USD')  # CCY
    row[14] = data.get('Strike Date', '')  # Strike Date
    row[15] = data.get('Issue Date', '')  # Issue Date
    row[16] = data.get('ISIN', '')  # ISIN
    # row[17] is empty
    
    # Underlying assets (18-21)
    underlyings = data.get('underlying_assets', [])
    for i in range(4):
        if i < len(underlyings):
            row[18 + i] = underlyings[i].get('Name', '')
        else:
            row[18 + i] = ''
    
    # row[22] is empty
    row[23] = data.get('Notional Value', '')  # Investment $
    row[24] = 1  # Total Units (simplified)
    row[25] = data.get('Notional Value', '')  # Notional Value
    row[26] = data.get('Notional Value', '') if data.get('CCY') == 'AUD' else ''  # AUD Equivalent
    row[27] = data.get('Maturity Date', '')  # Maturity Date
    row[28] = ''  # Revenue (to be calculated)
    # row[29] is empty
    
    # Underlying prices (30-43)
    for i in range(4):
        if i < len(underlyings):
            row[30 + i] = underlyings[i].get('Initial_Price', '')  # Issue Prices
            row[35 + i] = underlyings[i].get('Knock_In_Price', '')  # Knock In Prices  
            row[40 + i] = underlyings[i].get('Knock_Out_Price', '')  # Knock Out Prices
        else:
            row[30 + i] = ''
            row[35 + i] = ''
            row[40 + i] = ''
    
    return row

def append_to_fixed_income_master(new_row_data, master_file_path):
    """Append new data to your Fixed Income Desk Master File Database sheet"""
    
    try:
        # Read the existing file
        if not os.path.exists(master_file_path):
            return False, f"Master file not found: {master_file_path}"
        
        # Read all sheets to preserve them
        xl_file = pd.ExcelFile(master_file_path)
        
        # Read the Database sheet (note the space)
        database_df = pd.read_excel(master_file_path, sheet_name=' Database', header=0)
        
        # Add the new row
        new_row_df = pd.DataFrame([new_row_data])
        updated_df = pd.concat([database_df, new_row_df], ignore_index=True)
        
        # Write back to Excel, preserving all sheets
        with pd.ExcelWriter(master_file_path, engine='openpyxl', mode='w') as writer:
            # Write all existing sheets
            for sheet_name in xl_file.sheet_names:
                if sheet_name == ' Database':
                    updated_df.to_excel(writer, sheet_name=' Database', index=False, header=False)
                else:
                    existing_sheet = pd.read_excel(master_file_path, sheet_name=sheet_name)
                    existing_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
        
        return True, f"Successfully added row {len(updated_df)} to Database sheet"
        
    except Exception as e:
        return False, f"Error updating master file: {str(e)}"

# Initialize the extractor
extractor = FixedIncomeTermsheetExtractor()

# STREAMLIT APP
st.set_page_config(
    page_title="Fixed Income Desk Termsheet Extractor",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Fixed Income Desk Termsheet Extractor")
st.write("Upload PDF termsheets and append to your Fixed Income Desk Master File Database")

# File upload section
uploaded_files = st.file_uploader(
    "Fixed Income Desk Master File (Excel)",
    type=['xlsx'],
    help="Upload your Fixed_Income_Desk_Master_File.xlsx",
    key="master_file"
)

if uploaded_files:
    # Save the master file
    master_path = "/tmp/Fixed_Income_Desk_Master_File.xlsx"
    with open(master_path, "wb") as f:
        f.write(uploaded_files.read())
    st.success("✅ Master file loaded successfully!")
else:
    master_path = "/mnt/user-data/uploads/Fixed_Income_Desk_Master_File.xlsx"
    if os.path.exists(master_path):
        st.info("📁 Using uploaded Fixed Income Desk Master File")
    else:
        st.warning("⚠️ Please upload your Fixed_Income_Desk_Master_File.xlsx first")

# PDF termsheet uploader
pdf_file = st.file_uploader(
    "PDF Termsheet to Extract",
    type=['pdf'],
    help="Upload the PDF termsheet to extract data from"
)

if pdf_file is not None and os.path.exists(master_path):
    # Save uploaded PDF temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
        tmp_file.write(pdf_file.read())
        tmp_path = tmp_file.name
    
    if st.button("Extract and Add to Database", type="primary"):
        with st.spinner("Processing termsheet..."):
            try:
                # Extract data
                data = extractor.extract_termsheet_data(tmp_path)
                
                if 'error' not in data:
                    st.success(f"✅ Data extracted from {data.get('Detected_Issuer_Type', 'unknown').replace('_', ' ').title()}!")
                    
                    # Display preview
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.subheader("Core Information")
                        st.write(f"**Issuer:** {data.get('Issuer', 'N/A')}")
                        st.write(f"**ISIN:** {data.get('ISIN', 'N/A')}")
                        st.write(f"**Currency:** {data.get('CCY', 'N/A')}")
                        st.write(f"**Notional:** {data.get('Notional Value', 'N/A')}")
                    
                    with col2:
                        st.subheader("Key Dates")
                        st.write(f"**Issue Date:** {data.get('Issue Date', 'N/A')}")
                        st.write(f"**Strike Date:** {data.get('Strike Date', 'N/A')}")
                        st.write(f"**Maturity:** {data.get('Maturity Date', 'N/A')}")
                    
                    with col3:
                        st.subheader("Risk Parameters")
                        st.write(f"**Knock-In:** {data.get('Knock-In%', 'N/A')}")
                        st.write(f"**Coupon Rate:** {data.get('Coupon Rate - Annual', 'N/A')}")
                        st.write(f"**Underlyings:** {len(data.get('underlying_assets', []))}")
                    
                    # Create database row
                    database_row = create_database_row(data)
                    
                    # Show preview of what will be added
                    st.subheader("Preview of Database Row")
                    preview_dict = {}
                    for i, value in enumerate(database_row):
                        col_name = extractor.column_mapping.get(i, f"Column_{i}")
                        if col_name and value:
                            preview_dict[col_name] = value
                    
                    preview_df = pd.DataFrame([preview_dict])
                    st.dataframe(preview_df, use_container_width=True)
                    
                    # Append to master file
                    success, message = append_to_fixed_income_master(database_row, master_path)
                    
                    if success:
                        st.success(f"✅ {message}")
                        st.balloons()
                        
                        # Offer download of updated file
                        with open(master_path, "rb") as file:
                            st.download_button(
                                label="📥 Download Updated Master File",
                                data=file.read(),
                                file_name="Fixed_Income_Desk_Master_File_Updated.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error(f"❌ {message}")
                    
                    # Show underlying assets if found
                    if data.get('underlying_assets'):
                        st.subheader("Underlying Assets")
                        df_underlying = pd.DataFrame(data['underlying_assets'])
                        st.dataframe(df_underlying, use_container_width=True)
                
                else:
                    st.error(f"❌ Extraction failed: {data['error']}")
                    
            except Exception as e:
                st.error(f"❌ Processing error: {str(e)}")
        
        # Clean up temp file
        os.unlink(tmp_path)

# Show supported issuers
with st.expander("📋 Supported Issuers"):
    st.write("This extractor can handle termsheets from:")
    for issuer_key, config in extractor.issuer_patterns.items():
        if issuer_key != 'generic':
            st.write(f"• **{config['issuer_name']}** - {', '.join(config['identifiers'])}")
    st.write("• **Generic Pattern Matching** - For unlisted issuers")

# Show current database info
if os.path.exists(master_path):
    with st.expander("📊 Current Database Info"):
        try:
            current_df = pd.read_excel(master_path, sheet_name=' Database', header=0)
            st.write(f"Current rows in database: **{len(current_df)}**")
            if len(current_df) > 0:
                st.write("Recent entries:")
                display_df = current_df.tail(3).iloc[:, :10]  # Show last 3 rows, first 10 columns
                st.dataframe(display_df, use_container_width=True)
        except Exception as e:
            st.write(f"Could not read database: {e}")

st.markdown("---")
st.caption("Fixed Income Desk termsheet extraction with Database integration")
