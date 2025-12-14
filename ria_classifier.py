"""
RIA Data Multi-Category Classifier V7
Three-Prompt Approach: Separate prompts for Services, Clients, and Products
Each prompt extracts its primary fields PLUS all secondary fields
Results are merged and deduplicated
"""

import os
import json
import pandas as pd
import re
import time
from typing import Dict, List, Tuple, Optional, Any
from datetime import datetime
from openai import AzureOpenAI
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# ============================================================================
# CONFIGURATION
# ============================================================================

CONFIG = {
    "excel_path": "/Users/nikitakumar/Documents/RIA_Web_Scraping/WebScraping-Output/final_output/ria_data_multisheet.xlsx",
    "output_path": "/Users/nikitakumar/Documents/RIA_Web_Scraping/WebScraping-Output/final_output/ria_data_multisheet_final_v7.xlsx",
    "mappings_file": "/Users/nikitakumar/Documents/RIA_Web_Scraping/WebScraping-Output/final_output/final_standard_mappings_v7.json",
    "azure_endpoint": os.getenv("AZURE_OPENAI_ENDPOINT"),
    "azure_api_key": os.getenv("AZURE_OPENAI_API_KEY"),
    "azure_deployment": "gpt-5.1-chat",
    "azure_api_version": os.getenv("AZURE_OPENAI_API_VERSION"),
    "sheets_to_process": ["Services", "Target_Clients", "Products_Instruments"],
    "source_columns": {
        "Services": "services",
        "Target_Clients": "target_clients",
        "Products_Instruments": "products_instruments"
    },
    "max_retries": 2
}

# ============================================================================
# FIELD DEFINITIONS
# ============================================================================

# Primary fields for each prompt
PRIMARY_FIELDS = {
    "services": ["service_category"],
    "clients": ["client_category"],
    "products": ["asset_class", "asset_sub_class", "asset_subclass_further", "additional_assets"]
}

# Secondary fields - extracted by ALL prompts
SECONDARY_FIELDS = [
    "investment_strategy",
    "investment_theme",
    "investment_style",
    "return_objective",
    "asset_allocation_approach",
    "esg_values_category",
    "tax_management_approach",
    "risk_management_hedging",
    "specialized_strategy",
    "geographic_exposure",
    "market_capitalization",
    "duration_maturity",
    "credit_quality",
    "liquidity_profile",
    "management_style",
    "edge_cases"
]

# All output columns
OUTPUT_COLUMNS = (
    PRIMARY_FIELDS["services"] + 
    PRIMARY_FIELDS["clients"] + 
    PRIMARY_FIELDS["products"][:3] +  # Exclude additional_assets from main columns
    SECONDARY_FIELDS
)

# ============================================================================
# VALID VALUES FOR VALIDATION
# ============================================================================

VALID_VALUES = {
    "service_category": [
        "Discretionary Portfolio Management",
        "Non-Discretionary Portfolio Management",
        "Financial Planning",
        "Wealth Management",
        "Retirement / ERISA Plan Advisory",
        "Institutional Portfolio Management",
        "Individual / HNW Portfolio Management",
        "Wrap Fee / UMA Program Management",
        "Sub-Adviser Selection / Model Marketplace",
        "Robo-Advisory / Digital Advice",
        "Tax & Estate Planning",
        "Insurance & Annuity Advisory",
        "Other Services"
    ],
    
    "client_category": [
        "Individuals",
        "High-Net-Worth Individuals",
        "Family Offices",
        "Corporations / Businesses",
        "Banks & Financial Institutions",
        "Investment Companies",
        "Pooled Investment Vehicles",
        "Pension & Retirement Plans",
        "Government Entities",
        "Charitable Organizations / Nonprofits",
        "Insurance Companies / Separate Accounts",
        "Other Clients"
    ],
    
    "asset_class": [
        "Equity",
        "Fixed Income",
        "Cash & Cash Equivalents / Money Market",
        "Alternatives & Real Assets",
        "Derivatives & Overlay Strategies",
        "Insurance & Retirement Products",
        "Digital & Emerging Assets",
        "Pooled Vehicles / Wrappers",
        "Planning-Oriented / Household Assets"
    ],
    
    "asset_sub_class": [
        "U.S. Equity (Domestic)", "International & Global Equity", "Sector & Thematic Equity", "Specialized Equity",
        "Government & Municipal", "Corporate & Credit", "Securitized / Structured Credit", "Specialized Fixed Income",
        "Cash / Cash Reserves", "Money Market Funds", "Treasury Bills / Short-Term Government",
        "CDs (Certificates of Deposit)", "High-Yield Savings / Bank Deposit Programs",
        "Short-Term Investment Grade Instruments", "Sweep Vehicles / Bank Deposit Program (BDP)", "Commercial Paper",
        "Private Equity", "Private Credit & Distressed", "Real Estate", "Natural Resources & Infrastructure",
        "Commodities", "Hedge Funds & Absolute Return", "Specialized Real Assets",
        "Options Strategies", "Futures & Forwards", "Swaps & Credit Derivatives", "Risk-Management / Hedging Programs",
        "Annuities", "Life-Insurance-Linked Investments", "Institutional / Retirement Mandates",
        "Cryptocurrency / Digital Asset Exposure", "Blockchain / Web3 / Tokenization Exposure",
        "Mutual Funds", "ETFs", "Closed-End / Interval Funds", "SMA / UMA / Model Portfolios",
        "Private Funds / LPs / Hedge Funds (as vehicles)",
        "Retirement Accounts", "Tax-Advantaged Accounts", "Trust & Estate Accounts",
        "Concentrated / Employer Stock", "Charitable & Foundation Assets"
    ],
    
    "asset_subclass_further": [
        # Equity - U.S.
        "Large Cap", "Mid Cap", "Small Cap", "Micro Cap",
        "Large Cap Growth", "Large Cap Value", "Large Cap Blend",
        "Mid Cap Growth", "Mid Cap Value", "Mid Cap Blend",
        "Small Cap Growth", "Small Cap Value", "Small Cap Blend",
        "Dividend Stocks", "Blue Chip", "S&P 500", "Russell 2000", "Total US Market",
        # Equity - International
        "Developed Markets", "EAFE", "Europe", "Japan", "UK", "Asia Pacific", "Canada", "Australia",
        "Global Equity", "Global Ex-US", "World",
        # Equity - Sector
        "Technology", "Healthcare", "Financials", "Energy", "Consumer Discretionary",
        "Consumer Staples", "Industrials", "Materials", "Utilities", "Real Estate",
        "Communication Services", "Clean Energy", "AI & Robotics", "Cybersecurity", "Fintech",
        # Equity - Specialized
        "Emerging Markets Equity", "Frontier Markets Equity", "Single Country", "BRIC",
        "China", "India", "Brazil", "Master Limited Partnerships (MLPs)", "Preferred Stock",
        # Fixed Income
        "U.S. Treasuries", "Treasury Bills", "Treasury Notes", "Treasury Bonds", "TIPS", "I-Bonds",
        "Agency Bonds", "Ginnie Mae", "Fannie Mae", "Freddie Mac",
        "Municipal Bonds", "State Bonds", "Local Bonds", "Tax-Exempt Munis", "Taxable Munis",
        "General Obligation", "Revenue Bonds",
        "Investment Grade Corporate", "High Yield Corporate", "Junk Bonds",
        "Senior Secured", "Senior Unsecured", "Subordinated",
        "Floating Rate Notes", "Convertible Bonds", "Bank Loans", "Leveraged Loans", "Direct Lending",
        "Mortgage-Backed Securities (MBS)", "Agency MBS", "Non-Agency MBS",
        "Commercial MBS (CMBS)", "Residential MBS (RMBS)",
        "Asset-Backed Securities (ABS)", "CLOs", "CDOs",
        "Emerging Market Debt", "EM Hard Currency", "EM Local Currency",
        "Global Bonds", "International Bonds", "Inflation-Linked",
        # Alternatives
        "Buyout", "Growth Equity", "Venture Capital", "Early Stage VC", "Late Stage VC", "Seed",
        "Secondaries", "Co-Investments", "Fund of Funds PE",
        "Mezzanine", "Distressed Debt", "Special Situations",
        "Direct Real Estate", "Core Real Estate", "Core Plus", "Value-Add Real Estate",
        "Opportunistic Real Estate", "Public REITs", "Private REITs", "Mortgage REITs",
        "Commercial Real Estate", "Residential Real Estate", "Industrial Real Estate",
        "Infrastructure", "Core Infrastructure", "Timber", "Farmland", "Agriculture",
        "Gold", "Silver", "Precious Metals", "Oil", "Natural Gas",
        "Long/Short Equity HF", "Market Neutral HF", "Global Macro HF", "Event Driven HF",
        "Multi-Strategy HF", "CTA", "Managed Futures", "Fund of Hedge Funds",
        # Insurance & Retirement
        "Fixed Annuities", "Variable Annuities", "Indexed Annuities",
        "Immediate Annuities", "Deferred Annuities", "SPIA", "DIA", "QLAC",
        "Variable Universal Life (VUL)", "Indexed Universal Life (IUL)", "Whole Life",
        "Pension Fund", "Defined Benefit", "Defined Contribution", "ERISA Plans",
        # Pooled Vehicles
        "Passive ETF", "Active ETF", "Index ETF", "Thematic ETF", "Sector ETF",
        "Open-End Mutual Fund", "Index Mutual Fund", "Active Mutual Fund",
        "Separately Managed Account (SMA)", "Unified Managed Account (UMA)", "Model Portfolio",
        # Planning-Oriented
        "Traditional IRA", "Roth IRA", "SEP IRA", "SIMPLE IRA", "Solo 401(k)", "401(k)",
        "403(b)", "457(b)", "Rollover IRA", "Inherited IRA", "Backdoor Roth",
        "529 Plan", "Coverdell ESA", "HSA", "FSA", "UTMA", "UGMA",
        "Revocable Trust", "Irrevocable Trust", "Living Trust", "Testamentary Trust",
        "GRAT", "GRUT", "QPRT", "Dynasty Trust", "Spendthrift Trust", "Special Needs Trust",
        "Donor Advised Fund (DAF)", "Charitable Remainder Trust (CRT)",
        "Charitable Lead Trust (CLT)", "Private Foundation"
    ],
    
    "investment_strategy": [
        "Long/Short Equity", "Market Neutral", "Long Only", "Global Macro", "Event Driven",
        "Merger Arbitrage", "Distressed", "Activist", "Quantitative", "Growth", "Value",
        "GARP", "Momentum", "Factor Investing", "Smart Beta", "Index", "CTA / Managed Futures",
        "Multi-Strategy", "Absolute Return", "Relative Value", "Credit Long/Short",
        "Fixed Income Arbitrage", "Convertible Arbitrage", "Statistical Arbitrage",
        "Thematic Investing", "Sector Rotation", "Tactical Asset Allocation",
        "Strategic Asset Allocation", "Risk Parity", "Target Date", "Liability Driven",
        "Covered Call", "Protective Put", "Collar Strategy"
    ],
    
    "investment_theme": [
        "ESG / Sustainability", "SRI / Socially Responsible", "Impact Investing",
        "Green / Clean Energy", "Climate / Carbon", "Water", "Technology & Innovation",
        "AI / Machine Learning", "Robotics / Automation", "Cybersecurity",
        "Healthcare / Biotech", "Genomics", "Aging Demographics", "Millennials / Gen Z",
        "Emerging Markets", "Frontier Markets", "China", "India", "Southeast Asia",
        "Infrastructure", "Smart Cities", "Digital Assets / Blockchain", "Fintech",
        "E-commerce", "Cloud Computing", "5G / Connectivity", "Space / Aerospace",
        "Defense / Security", "Consumer Trends"
    ],
    
    "investment_style": ["Growth", "Value", "Core / Blend", "GARP", "Momentum", "Contrarian"],
    
    "return_objective": [
        "Income / Yield-Focused", "Total Return", "Capital Preservation",
        "Capital Appreciation", "Absolute Return"
    ],
    
    "asset_allocation_approach": [
        "Balanced / 60/40", "Strategic Asset Allocation",
        "Tactical Asset Allocation / Dynamic Allocation", "Multi-Asset / Diversified",
        "Risk Parity", "Portable Alpha", "Factor / Smart Beta"
    ],
    
    "esg_values_category": [
        "ESG Integration", "Sustainable Investing", "Impact Investing",
        "SRI / Socially Responsible Investing", "Faith-Based / Values-Based",
        "Negative Screening", "Positive Screening", "ESG Thematic"
    ],
    
    "tax_management_approach": [
        "Tax-Managed / Tax-Efficient", "Tax-Loss Harvesting",
        "Municipal Bond Focus (Tax-Exempt)", "After-Tax Alpha Focus"
    ],
    
    "risk_management_hedging": [
        "Options Overlay / Hedged Equity", "Covered Call Strategies",
        "Downside Protection / Tail Risk Hedging", "Currency Hedged", "Market Neutral"
    ],
    
    "specialized_strategy": [
        "Liability-Driven Investing (LDI)", "Outcome-Oriented / Goals-Based",
        "OCIO / Outsourced CIO", "Endowment-Style / Yale Model",
        "All-Weather / Risk Parity", "Defined Outcome / Buffer Strategies"
    ],
    
    "geographic_exposure": [
        "U.S. / Domestic", "International Developed", "Emerging Markets",
        "Frontier Markets", "Global", "Regional", "Country-Specific"
    ],
    
    "market_capitalization": ["Large Cap", "Mid Cap", "Small Cap", "Micro Cap", "All Cap / Multi Cap"],
    
    "duration_maturity": [
        "Ultra-Short (<1 year)", "Short-Term (1-3 years)",
        "Intermediate-Term (3-10 years)", "Long-Term (10+ years)", "Target Maturity"
    ],
    
    "credit_quality": [
        "AAA/AA", "A/BBB (Investment Grade)", "BB/B (High Yield)",
        "CCC and Below (Distressed)", "Unrated"
    ],
    
    "liquidity_profile": [
        "Daily Liquidity", "Monthly / Quarterly Liquidity",
        "Limited Liquidity", "Illiquid / Lock-Up"
    ],
    
    "management_style": [
        "Active Management", "Passive / Index", "Smart Beta / Factor-Based",
        "Quantitative / Systematic", "Fundamental", "Rules-Based"
    ],
    
    "edge_cases": [
        "Custom Solutions / Bespoke Strategies", "Alternative Investments",
        "Other Investments (as disclosed in Form ADV)", "Proprietary Strategies",
        "Multi-Strategy / Multi-Asset (when specific allocation not disclosed)",
        "To Be Determined / Under Review"
    ]
}

# ============================================================================
# SECONDARY FIELDS TAXONOMY (Used in all 3 prompts)
# ============================================================================

SECONDARY_FIELDS_TAXONOMY = """
## SECONDARY FIELDS - Extract if mentioned in the data:

### investment_strategy
Long/Short Equity | Market Neutral | Long Only | Global Macro | Event Driven | Merger Arbitrage | Distressed | Activist | Quantitative | Growth | Value | GARP | Momentum | Factor Investing | Smart Beta | Index | CTA / Managed Futures | Multi-Strategy | Absolute Return | Relative Value | Thematic Investing | Sector Rotation | Tactical Asset Allocation | Strategic Asset Allocation | Risk Parity | Target Date | Liability Driven | Covered Call | Protective Put | Collar Strategy

### investment_theme
ESG / Sustainability | SRI / Socially Responsible | Impact Investing | Green / Clean Energy | Climate / Carbon | Water | Technology & Innovation | AI / Machine Learning | Robotics / Automation | Cybersecurity | Healthcare / Biotech | Genomics | Aging Demographics | Infrastructure | Digital Assets / Blockchain | Fintech | Consumer Trends

### investment_style
Growth | Value | Core / Blend | GARP | Momentum | Contrarian

### return_objective
Income / Yield-Focused | Total Return | Capital Preservation | Capital Appreciation | Absolute Return

### asset_allocation_approach
Balanced / 60/40 | Strategic Asset Allocation | Tactical Asset Allocation / Dynamic Allocation | Multi-Asset / Diversified | Risk Parity | Portable Alpha | Factor / Smart Beta

### esg_values_category
ESG Integration | Sustainable Investing | Impact Investing | SRI / Socially Responsible Investing | Faith-Based / Values-Based | Negative Screening | Positive Screening | ESG Thematic

### tax_management_approach
| If you see... | Map to |
|---------------|--------|
| "Tax Planning", "Tax Optimization", "Tax-Efficient", "Tax Strategies" | Tax-Managed / Tax-Efficient |
| "Tax-Loss Harvesting" | Tax-Loss Harvesting |
| "Municipal", "Muni", "Tax-Exempt", "Municipal Bond" | Municipal Bond Focus (Tax-Exempt) |

**ALLOWED:** Tax-Managed / Tax-Efficient | Tax-Loss Harvesting | Municipal Bond Focus (Tax-Exempt) | After-Tax Alpha Focus

### risk_management_hedging
Options Overlay / Hedged Equity | Covered Call Strategies | Downside Protection / Tail Risk Hedging | Currency Hedged | Market Neutral

### specialized_strategy
| If you see... | Map to |
|---------------|--------|
| "Goals-Based", "Outcome-Oriented", "Comprehensive Advisory" | Outcome-Oriented / Goals-Based |
| "OCIO", "Outsourced CIO" | OCIO / Outsourced CIO |
| "Endowment", "Yale Model" | Endowment-Style / Yale Model |
| "Buffer", "Defined Outcome" | Defined Outcome / Buffer Strategies |

**ALLOWED:** Liability-Driven Investing (LDI) | Outcome-Oriented / Goals-Based | OCIO / Outsourced CIO | Endowment-Style / Yale Model | All-Weather / Risk Parity | Defined Outcome / Buffer Strategies

### geographic_exposure
| If you see... | Map to |
|---------------|--------|
| "Global", "Global Equity", "Worldwide" | Global |
| "Non-US", "Non-US Equity", "International", "Foreign" | International Developed |
| "Domestic", "US", "U.S.", "United States" | U.S. / Domestic |
| "Emerging Markets", "EM", "Developing" | Emerging Markets |
| "EAFE", "Developed Markets" | International Developed |

**ALLOWED:** U.S. / Domestic | International Developed | Emerging Markets | Frontier Markets | Global | Regional | Country-Specific

### market_capitalization
| If you see... | Map to |
|---------------|--------|
| "Large Cap", "Large-Cap", "Blue Chip", "Mega Cap" | Large Cap |
| "Mid Cap", "Mid-Cap", "Medium Cap" | Mid Cap |
| "Small Cap", "Small-Cap", "Small Cap Platform" | Small Cap |
| "Micro Cap" | Micro Cap |
| "All Cap", "Multi Cap", "Full Cap" | All Cap / Multi Cap |

**ALLOWED:** Large Cap | Mid Cap | Small Cap | Micro Cap | All Cap / Multi Cap

### duration_maturity
Ultra-Short (<1 year) | Short-Term (1-3 years) | Intermediate-Term (3-10 years) | Long-Term (10+ years) | Target Maturity

### credit_quality
AAA/AA | A/BBB (Investment Grade) | BB/B (High Yield) | CCC and Below (Distressed) | Unrated

### liquidity_profile
| If you see... | Map to |
|---------------|--------|
| "Daily", "Liquid", "Public" | Daily Liquidity |
| "Monthly", "Quarterly" | Monthly / Quarterly Liquidity |
| "Limited", "Semi-Liquid" | Limited Liquidity |
| "Illiquid", "Lock-Up", "Private Equity", "Private" | Illiquid / Lock-Up |

**ALLOWED:** Daily Liquidity | Monthly / Quarterly Liquidity | Limited Liquidity | Illiquid / Lock-Up

### management_style
| If you see... | Map to |
|---------------|--------|
| "Active", "Actively Managed" | Active Management |
| "Passive", "Index", "Indexed", "Indexing" | Passive / Index |
| "Smart Beta", "Factor" | Smart Beta / Factor-Based |
| "Quantitative", "Systematic", "Quant" | Quantitative / Systematic |

**ALLOWED:** Active Management | Passive / Index | Smart Beta / Factor-Based | Quantitative / Systematic | Fundamental | Rules-Based

### edge_cases
| If you see... | Map to |
|---------------|--------|
| "Alternative Investments", "Alternatives", "Alternative" | Alternative Investments |
| "Custom", "Customized", "Bespoke", "Tailored" | Custom Solutions / Bespoke Strategies |
| "Proprietary" | Proprietary Strategies |
| "Multi-Strategy", "Multi-Asset" | Multi-Strategy / Multi-Asset (when specific allocation not disclosed) |

**ALLOWED:** Custom Solutions / Bespoke Strategies | Alternative Investments | Other Investments (as disclosed in Form ADV) | Proprietary Strategies | Multi-Strategy / Multi-Asset (when specific allocation not disclosed) | To Be Determined / Under Review
"""


# ============================================================================
# PROMPT 1: SERVICES CLASSIFIER
# ============================================================================

def build_services_prompt(crd: str, services_raw: str) -> str:
    """Build prompt for classifying services data."""
    
    prompt = f"""# SERVICES CLASSIFIER

You are classifying SERVICES data for investment advisor CRD {crd}.

## PRIMARY FIELD TO EXTRACT: service_category

### MAPPING TABLE:
| If you see... | Map to service_category |
|---------------|------------------------|
| "Wealth Management", "Wealth Advisory", "Private Wealth", "Wealth Strategy" | Wealth Management |
| "Financial Planning", "Comprehensive Planning", "Retirement Planning" | Financial Planning |
| "Estate Planning", "Trust Services", "Trustee", "Trustee Services" | Tax & Estate Planning |
| "Tax Planning", "Tax Strategies", "Tax Advisory", "Income Tax Planning" | Tax & Estate Planning |
| "Insurance Advisory", "Insurance Solutions", "Annuity Solutions", "Insurance and Annuity" | Insurance & Annuity Advisory |
| "Investment Advisory", "Investment Management", "Asset Management" | Discretionary Portfolio Management |
| "Portfolio Management", "Managed Accounts", "Investment portfolios" | Discretionary Portfolio Management |
| "Retirement Planning", "401k", "ERISA", "Pension", "Retirement Plan Products" | Retirement / ERISA Plan Advisory |
| "Institutional", "Institutional Advisory", "Institutional investors" | Institutional Portfolio Management |
| "Family Office", "Family Office Solutions", "Global Family Office" | Wealth Management |
| "Philanthropic Advisory", "Foundation Administration", "Philanthropy" | Financial Planning |
| "HNW", "High Net Worth", "UHNW" | Individual / HNW Portfolio Management |
| "Wrap Fee", "UMA", "Unified Managed" | Wrap Fee / UMA Program Management |
| "Robo", "Digital Advice", "Automated" | Robo-Advisory / Digital Advice |

### ALLOWED service_category VALUES:
Discretionary Portfolio Management | Non-Discretionary Portfolio Management | Financial Planning | Wealth Management | Retirement / ERISA Plan Advisory | Institutional Portfolio Management | Individual / HNW Portfolio Management | Wrap Fee / UMA Program Management | Sub-Adviser Selection / Model Marketplace | Robo-Advisory / Digital Advice | Tax & Estate Planning | Insurance & Annuity Advisory | Other Services

---

{SECONDARY_FIELDS_TAXONOMY}

---

## DATA TO CLASSIFY:

CRD: {crd}
SERVICES DATA:
{services_raw}

---

## OUTPUT FORMAT (JSON only, no markdown):

{{
  "service_category": "use | for multiple values",
  "investment_strategy": "",
  "investment_theme": "",
  "investment_style": "",
  "return_objective": "",
  "asset_allocation_approach": "",
  "esg_values_category": "",
  "tax_management_approach": "",
  "risk_management_hedging": "",
  "specialized_strategy": "",
  "geographic_exposure": "",
  "market_capitalization": "",
  "duration_maturity": "",
  "credit_quality": "",
  "liquidity_profile": "",
  "management_style": "",
  "edge_cases": ""
}}

⚠️ RULES:
- Use EXACT values from taxonomy only
- Use " | " to separate multiple values
- If not found, use empty string ""
- Return ONLY valid JSON
"""
    return prompt


# ============================================================================
# PROMPT 2: CLIENTS CLASSIFIER
# ============================================================================

def build_clients_prompt(crd: str, clients_raw: str) -> str:
    """Build prompt for classifying clients data."""
    
    prompt = f"""# CLIENTS CLASSIFIER

You are classifying CLIENTS/TARGET AUDIENCE data for investment advisor CRD {crd}.

## PRIMARY FIELD TO EXTRACT: client_category

### MAPPING TABLE:
| If you see... | Map to client_category |
|---------------|----------------------|
| "HNW", "High Net Worth", "High-Net-Worth", "HNWI" | High-Net-Worth Individuals |
| "UHNW", "Ultra High Net Worth", "Ultra-High-Net-Worth" | High-Net-Worth Individuals |
| "Family Office", "Multi-Family Office", "Single Family Office", "Family Offices" | Family Offices |
| "Institution", "Institutional", "Institutional investors", "Institutions" | Investment Companies |
| "Corporation", "Corporate", "Corporations", "Businesses", "Business Owners" | Corporations / Businesses |
| "Individual", "Individuals", "Families", "Retail" | Individuals |
| "Foundation", "Endowment", "Non-profit", "Charitable", "Nonprofits" | Charitable Organizations / Nonprofits |
| "Pension", "Retirement Plan", "401k Plan", "ERISA", "Retirees" | Pension & Retirement Plans |
| "Bank", "Financial Institution", "Banks" | Banks & Financial Institutions |
| "Government", "Municipal", "State", "Federal" | Government Entities |
| "Insurance Company" | Insurance Companies / Separate Accounts |
| "Pooled Vehicle", "Fund", "LP" | Pooled Investment Vehicles |
| "Entrepreneur", "Executive", "Senior Executives", "Executives" | High-Net-Worth Individuals |
| "Fortune 1000", "Fortune 1000 Employees" | Corporations / Businesses |
| "Professional athletes", "Athletes" | High-Net-Worth Individuals |
| "Visionaries", "Innovators" | Other Clients |

### ALLOWED client_category VALUES:
Individuals | High-Net-Worth Individuals | Family Offices | Corporations / Businesses | Banks & Financial Institutions | Investment Companies | Pooled Investment Vehicles | Pension & Retirement Plans | Government Entities | Charitable Organizations / Nonprofits | Insurance Companies / Separate Accounts | Other Clients

---

{SECONDARY_FIELDS_TAXONOMY}

---

## DATA TO CLASSIFY:

CRD: {crd}
CLIENTS DATA:
{clients_raw}

---

## OUTPUT FORMAT (JSON only, no markdown):

{{
  "client_category": "use | for multiple values",
  "investment_strategy": "",
  "investment_theme": "",
  "investment_style": "",
  "return_objective": "",
  "asset_allocation_approach": "",
  "esg_values_category": "",
  "tax_management_approach": "",
  "risk_management_hedging": "",
  "specialized_strategy": "",
  "geographic_exposure": "",
  "market_capitalization": "",
  "duration_maturity": "",
  "credit_quality": "",
  "liquidity_profile": "",
  "management_style": "",
  "edge_cases": ""
}}

⚠️ RULES:
- Use EXACT values from taxonomy only
- Use " | " to separate multiple values
- If not found, use empty string ""
- Return ONLY valid JSON
"""
    return prompt


# ============================================================================
# PROMPT 3: PRODUCTS/ASSETS CLASSIFIER
# ============================================================================

def build_products_prompt(crd: str, products_raw: str) -> str:
    """Build prompt for classifying products/instruments data."""
    
    prompt = f"""# PRODUCTS/ASSETS CLASSIFIER

You are classifying PRODUCTS and INVESTMENT INSTRUMENTS data for investment advisor CRD {crd}.

## PRIMARY FIELDS TO EXTRACT: asset_class, asset_sub_class, asset_subclass_further, additional_assets

---

## HIERARCHICAL ASSET TAXONOMY

### ASSET_CLASS: "Equity"
**ASSET_SUB_CLASS options:**
- "U.S. Equity (Domestic)" → FURTHER: Large Cap | Mid Cap | Small Cap | Micro Cap | Large Cap Growth | Large Cap Value | Small Cap Growth | Small Cap Value | Dividend Stocks | Blue Chip | S&P 500 | Russell 2000
- "International & Global Equity" → FURTHER: Developed Markets | EAFE | Europe | Japan | UK | Asia Pacific | Global Equity | Global Ex-US | World
- "Sector & Thematic Equity" → FURTHER: Technology | Healthcare | Financials | Energy | Consumer Discretionary | Industrials | Clean Energy | AI & Robotics | Cybersecurity | Fintech
- "Specialized Equity" → FURTHER: Emerging Markets Equity | Frontier Markets Equity | China | India | Brazil | Preferred Stock

**MAPPING:**
| If you see... | Map to |
|---------------|--------|
| "Small Cap", "Small Cap Platform", "Small Cap Investments" | asset_class: Equity, asset_sub_class: U.S. Equity (Domestic), asset_subclass_further: Small Cap |
| "Large Cap", "Blue Chip" | asset_class: Equity, asset_sub_class: U.S. Equity (Domestic), asset_subclass_further: Large Cap |
| "Global Equity", "Non-US Equity", "International Equity" | asset_class: Equity, asset_sub_class: International & Global Equity, asset_subclass_further: Global Equity |
| "Thematic", "Thematic Platform", "Thematic Investments" | asset_class: Equity, asset_sub_class: Sector & Thematic Equity |

---

### ASSET_CLASS: "Fixed Income"
**ASSET_SUB_CLASS options:**
- "Government & Municipal" → FURTHER: U.S. Treasuries | Treasury Bills | Treasury Bonds | TIPS | Municipal Bonds | Tax-Exempt Munis | General Obligation | Revenue Bonds
- "Corporate & Credit" → FURTHER: Investment Grade Corporate | High Yield Corporate | Bank Loans | Leveraged Loans | Direct Lending
- "Securitized / Structured Credit" → FURTHER: Mortgage-Backed Securities (MBS) | Asset-Backed Securities (ABS) | CLOs
- "Specialized Fixed Income" → FURTHER: Emerging Market Debt | Global Bonds | International Bonds

**MAPPING:**
| If you see... | Map to |
|---------------|--------|
| "Municipal bonds", "Muni", "Municipal Bond ETF" | asset_class: Fixed Income, asset_sub_class: Government & Municipal, asset_subclass_further: Municipal Bonds |
| "Fixed Income", "Fixed Income Platform" | asset_class: Fixed Income |
| "Bonds", "Bond" | asset_class: Fixed Income |

---

### ASSET_CLASS: "Alternatives & Real Assets"
**ASSET_SUB_CLASS options:**
- "Private Equity" → FURTHER: Buyout | Growth Equity | Venture Capital | Early Stage VC | Late Stage VC | Seed | Secondaries | Co-Investments
- "Private Credit & Distressed" → FURTHER: Direct Lending | Mezzanine | Distressed Debt | Special Situations
- "Real Estate" → FURTHER: Direct Real Estate | Core Real Estate | Public REITs | Private REITs | Commercial Real Estate | Residential Real Estate
- "Natural Resources & Infrastructure" → FURTHER: Infrastructure | Timber | Farmland | Agriculture | Renewable Energy
- "Commodities" → FURTHER: Gold | Silver | Precious Metals | Oil | Natural Gas
- "Hedge Funds & Absolute Return" → FURTHER: Long/Short Equity HF | Market Neutral HF | Global Macro HF | Multi-Strategy HF | CTA | Managed Futures

**MAPPING:**
| If you see... | Map to |
|---------------|--------|
| "Private Equity", "Private Equity Partners", "Private investment opportunities" | asset_class: Alternatives & Real Assets, asset_sub_class: Private Equity |
| "Alternative Investments", "Alternatives", "Alternative Investments Platform" | asset_class: Alternatives & Real Assets |
| "Hedge Fund" | asset_class: Alternatives & Real Assets, asset_sub_class: Hedge Funds & Absolute Return |
| "Real Estate", "REITs" | asset_class: Alternatives & Real Assets, asset_sub_class: Real Estate |

---

### ASSET_CLASS: "Insurance & Retirement Products"
**ASSET_SUB_CLASS options:**
- "Annuities" → FURTHER: Fixed Annuities | Variable Annuities | Indexed Annuities | Immediate Annuities | Deferred Annuities
- "Life-Insurance-Linked Investments" → FURTHER: Variable Universal Life (VUL) | Indexed Universal Life (IUL) | Whole Life
- "Institutional / Retirement Mandates" → FURTHER: Pension Fund | Defined Benefit | Defined Contribution | ERISA Plans | 403(b) | 457 Plans

**MAPPING:**
| If you see... | Map to |
|---------------|--------|
| "Annuities", "Annuity solutions", "Insurance and Annuities" | asset_class: Insurance & Retirement Products, asset_sub_class: Annuities |
| "Insurance Products", "Insurance solutions" | asset_class: Insurance & Retirement Products |
| "Retirement Plan Products", "Retirement plan products" | asset_class: Insurance & Retirement Products, asset_sub_class: Institutional / Retirement Mandates |

---

### ASSET_CLASS: "Pooled Vehicles / Wrappers"
**ASSET_SUB_CLASS options:**
- "Mutual Funds" → FURTHER: Open-End Mutual Fund | Index Mutual Fund | Active Mutual Fund
- "ETFs" → FURTHER: Passive ETF | Active ETF | Index ETF | Thematic ETF | Sector ETF | Buffer ETF
- "Closed-End / Interval Funds" → FURTHER: Closed-End Fund | Interval Fund | BDC
- "SMA / UMA / Model Portfolios" → FURTHER: Separately Managed Account (SMA) | Unified Managed Account (UMA) | Model Portfolio

**MAPPING:**
| If you see... | Map to |
|---------------|--------|
| "ETF", "ETFs", "Municipal Bond ETF" | asset_class: Pooled Vehicles / Wrappers, asset_sub_class: ETFs |
| "Mutual Fund" | asset_class: Pooled Vehicles / Wrappers, asset_sub_class: Mutual Funds |

---

### ASSET_CLASS: "Planning-Oriented / Household Assets"
**ASSET_SUB_CLASS options:**
- "Retirement Accounts" → FURTHER: Traditional IRA | Roth IRA | SEP IRA | SIMPLE IRA | Solo 401(k) | 401(k) | 403(b) | 457(b) | Rollover IRA
- "Tax-Advantaged Accounts" → FURTHER: 529 Plan | Coverdell ESA | HSA | FSA | UTMA | UGMA
- "Trust & Estate Accounts" → FURTHER: Revocable Trust | Irrevocable Trust | Living Trust | Dynasty Trust | Special Needs Trust
- "Charitable & Foundation Assets" → FURTHER: Donor Advised Fund (DAF) | Charitable Remainder Trust (CRT) | Private Foundation

**MAPPING:**
| If you see... | Map to |
|---------------|--------|
| "Trusts", "Trust Solutions", "Trust services" | asset_class: Planning-Oriented / Household Assets, asset_sub_class: Trust & Estate Accounts |
| "Dynasty Trust", "Dynasty Trusts" | asset_class: Planning-Oriented / Household Assets, asset_sub_class: Trust & Estate Accounts, asset_subclass_further: Dynasty Trust |
| "Revocable Trust", "Revocable Trusts" | asset_class: Planning-Oriented / Household Assets, asset_sub_class: Trust & Estate Accounts, asset_subclass_further: Revocable Trust |
| "Annual Exclusion Gift Trusts" | asset_class: Planning-Oriented / Household Assets, asset_sub_class: Trust & Estate Accounts |
| "Retirement accounts", "IRA", "401k" | asset_class: Planning-Oriented / Household Assets, asset_sub_class: Retirement Accounts |
| "529", "529 Plan" | asset_class: Planning-Oriented / Household Assets, asset_sub_class: Tax-Advantaged Accounts, asset_subclass_further: 529 Plan |

---

### ASSET_CLASS: "Cash & Cash Equivalents / Money Market"
- "Traditional Investments" with no specifics → Consider: Equity | Fixed Income | Cash & Cash Equivalents / Money Market

---

{SECONDARY_FIELDS_TAXONOMY}

---

## DATA TO CLASSIFY:

CRD: {crd}
PRODUCTS/INSTRUMENTS DATA:
{products_raw}

---

## OUTPUT FORMAT (JSON only, no markdown):

{{
  "asset_class": "primary asset class (use | for multiple)",
  "asset_sub_class": "sub-classes (use | for multiple)",
  "asset_subclass_further": "most specific levels (use | for multiple)",
  "additional_assets": [
    {{"asset_class": "", "asset_sub_class": "", "asset_subclass_further": ""}}
  ],
  "investment_strategy": "",
  "investment_theme": "",
  "investment_style": "",
  "return_objective": "",
  "asset_allocation_approach": "",
  "esg_values_category": "",
  "tax_management_approach": "",
  "risk_management_hedging": "",
  "specialized_strategy": "",
  "geographic_exposure": "",
  "market_capitalization": "",
  "duration_maturity": "",
  "credit_quality": "",
  "liquidity_profile": "",
  "management_style": "",
  "edge_cases": ""
}}

⚠️ CRITICAL RULES:
1. Extract ALL asset classes mentioned - use additional_assets for secondary ones
2. For each asset, try to identify all 3 levels (class → sub_class → further)
3. "Private Equity" → asset_sub_class: Private Equity (under Alternatives & Real Assets)
4. "ETF", "Municipal Bond ETF" → asset_sub_class: ETFs (under Pooled Vehicles / Wrappers)
5. "Dynasty Trust", "Revocable Trust" → asset_subclass_further level
6. Use EXACT values from taxonomy only
7. Return ONLY valid JSON
"""
    return prompt


# ============================================================================
# LLM API CALL
# ============================================================================

def call_llm_api(
    client: AzureOpenAI,
    deployment: str,
    prompt: str,
    prompt_type: str,
    max_retries: int = 2
) -> Dict:
    """Call Azure OpenAI API with a specific prompt."""
    
    system_prompt = f"""You are a financial services categorization expert specializing in {prompt_type}.
Your task is to classify investment advisor data into predefined categories.

CRITICAL RULES:
1. Use ONLY exact values from the provided taxonomy
2. DO NOT create, combine, or modify category names
3. If a value doesn't match any category exactly, use empty string ""
4. For multiple values in one field, separate with " | "
5. Return ONLY valid JSON, no markdown, no explanation"""

    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=deployment,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt}
                ],
           
                response_format={"type": "json_object"}
            )
            
            content = response.choices[0].message.content.strip()
            content = re.sub(r"```json\s*", "", content)
            content = re.sub(r"```\s*", "", content)
            
            result = json.loads(content)
            return result
            
        except json.JSONDecodeError as e:
            print(f"      ⚠️ JSON Error ({prompt_type}, attempt {attempt + 1}): {e}")
            if attempt < max_retries - 1:
                time.sleep(1)
        except Exception as e:
            print(f"      ⚠️ API Error ({prompt_type}, attempt {attempt + 1}): {e}")
            if attempt < max_retries - 1:
                time.sleep(1)
    
    return {}


# ============================================================================
# MERGE RESULTS FROM 3 PROMPTS
# ============================================================================

def merge_values(values: List[str]) -> str:
    """Merge multiple values, deduplicate, and return pipe-separated string."""
    all_values = []
    for v in values:
        if v and v.strip():
            # Split by pipe if already pipe-separated
            parts = [p.strip() for p in v.split(" | ") if p.strip()]
            all_values.extend(parts)
    
    # Deduplicate while preserving order
    seen = set()
    unique = []
    for val in all_values:
        if val not in seen:
            seen.add(val)
            unique.append(val)
    
    return " | ".join(unique)


def merge_classification_results(
    services_result: Dict,
    clients_result: Dict,
    products_result: Dict
) -> Dict:
    """Merge results from all 3 prompts into a single classification."""
    
    merged = {}
    
    # Primary fields - from their specific prompts
    merged["service_category"] = services_result.get("service_category", "")
    merged["client_category"] = clients_result.get("client_category", "")
    merged["asset_class"] = products_result.get("asset_class", "")
    merged["asset_sub_class"] = products_result.get("asset_sub_class", "")
    merged["asset_subclass_further"] = products_result.get("asset_subclass_further", "")
    merged["additional_assets"] = products_result.get("additional_assets", [])
    
    # Secondary fields - merge from all 3 sources
    for field in SECONDARY_FIELDS:
        values = [
            services_result.get(field, ""),
            clients_result.get(field, ""),
            products_result.get(field, "")
        ]
        merged[field] = merge_values(values)
    
    return merged


# ============================================================================
# CLASSIFY SINGLE CRD
# ============================================================================

def classify_crd(
    client: AzureOpenAI,
    deployment: str,
    crd: str,
    services_raw: str,
    clients_raw: str,
    products_raw: str,
    max_retries: int = 2
) -> Dict:
    """
    Classify a single CRD using 3 specialized prompts.
    Returns merged results.
    """
    
    # Prompt 1: Services
    if services_raw and services_raw.strip():
        services_prompt = build_services_prompt(crd, services_raw)
        services_result = call_llm_api(client, deployment, services_prompt, "services", max_retries)
    else:
        services_result = {}
    
    # Prompt 2: Clients
    if clients_raw and clients_raw.strip():
        clients_prompt = build_clients_prompt(crd, clients_raw)
        clients_result = call_llm_api(client, deployment, clients_prompt, "clients", max_retries)
    else:
        clients_result = {}
    
    # Prompt 3: Products
    if products_raw and products_raw.strip():
        products_prompt = build_products_prompt(crd, products_raw)
        products_result = call_llm_api(client, deployment, products_prompt, "products", max_retries)
    else:
        products_result = {}
    
    # Merge results
    merged = merge_classification_results(services_result, clients_result, products_result)
    
    return merged


# ============================================================================
# VALIDATION
# ============================================================================

def validate_single_value(field: str, value: str) -> Tuple[bool, str]:
    """Validate a single field value against allowed values."""
    if not value or not value.strip():
        return True, ""
    
    value = value.strip()
    
    if field not in VALID_VALUES:
        return True, value
    
    valid_list = VALID_VALUES[field]
    
    # Handle pipe-separated multiple values
    if " | " in value:
        parts = [p.strip() for p in value.split(" | ")]
        validated_parts = []
        
        for part in parts:
            is_valid, corrected = validate_single_value(field, part)
            if is_valid and corrected:
                validated_parts.append(corrected)
        
        return True, " | ".join(validated_parts)
    
    # Exact match
    if value in valid_list:
        return True, value
    
    # Case-insensitive match
    value_lower = value.lower()
    for valid_val in valid_list:
        if value_lower == valid_val.lower():
            return True, valid_val
    
    # Normalized match
    def normalize(s):
        return s.lower().replace(" ", "").replace("-", "").replace("/", "")
    
    value_norm = normalize(value)
    for valid_val in valid_list:
        if value_norm == normalize(valid_val):
            return True, valid_val
    
    return False, ""


def validate_classification(crd: str, classification: Dict) -> Tuple[Dict, List[str]]:
    """Validate all fields in a classification."""
    validated = {}
    invalid_fields = []
    
    all_fields = OUTPUT_COLUMNS + ["additional_assets"]
    
    for field in all_fields:
        if field == "additional_assets":
            validated[field] = classification.get(field, [])
            continue
            
        value = classification.get(field, "")
        is_valid, corrected = validate_single_value(field, value)
        
        if is_valid:
            validated[field] = corrected
        else:
            validated[field] = ""
            invalid_fields.append(f"{field}='{value}'")
    
    return validated, invalid_fields


# ============================================================================
# MERGE ADDITIONAL ASSETS INTO MAIN COLUMNS
# ============================================================================

def merge_additional_assets_into_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Extract asset_class, asset_sub_class, asset_subclass_further from additional_assets
    and merge them into the main columns.
    """
    df = df.copy()
    
    for idx, row in df.iterrows():
        additional_assets_str = row.get("additional_assets", "[]")
        
        # Parse JSON
        try:
            if pd.isna(additional_assets_str) or additional_assets_str == "":
                additional_assets = []
            else:
                additional_assets = json.loads(additional_assets_str)
        except (json.JSONDecodeError, TypeError):
            additional_assets = []
        
        if not additional_assets or not isinstance(additional_assets, list):
            continue
        
        # Extract values from additional_assets
        additional_classes = []
        additional_sub_classes = []
        additional_further = []
        
        for asset in additional_assets:
            if isinstance(asset, dict):
                ac = asset.get("asset_class", "")
                asc = asset.get("asset_sub_class", "")
                asf = asset.get("asset_subclass_further", "")
                
                if ac and ac.strip():
                    additional_classes.append(ac.strip())
                if asc and asc.strip():
                    additional_sub_classes.append(asc.strip())
                if asf and asf.strip():
                    additional_further.append(asf.strip())
        
        # Get existing values
        existing_class = str(row.get("asset_class", "")) if pd.notna(row.get("asset_class")) else ""
        existing_sub = str(row.get("asset_sub_class", "")) if pd.notna(row.get("asset_sub_class")) else ""
        existing_further = str(row.get("asset_subclass_further", "")) if pd.notna(row.get("asset_subclass_further")) else ""
        
        # Parse existing pipe-separated values
        existing_classes = [v.strip() for v in existing_class.split(" | ") if v.strip()]
        existing_subs = [v.strip() for v in existing_sub.split(" | ") if v.strip()]
        existing_furthers = [v.strip() for v in existing_further.split(" | ") if v.strip()]
        
        # Merge and deduplicate
        def merge_unique(existing: list, additional: list) -> str:
            combined = existing + additional
            seen = set()
            unique = []
            for val in combined:
                if val not in seen:
                    seen.add(val)
                    unique.append(val)
            return " | ".join(unique)
        
        # Update dataframe
        df.at[idx, "asset_class"] = merge_unique(existing_classes, additional_classes)
        df.at[idx, "asset_sub_class"] = merge_unique(existing_subs, additional_sub_classes)
        df.at[idx, "asset_subclass_further"] = merge_unique(existing_furthers, additional_further)
    
    return df

# ============================================================================
# DATA LOADING FUNCTIONS
# ============================================================================

def load_excel_file(excel_path: str) -> Dict[str, pd.DataFrame]:
    """Load all sheets from an Excel file."""
    print(f"\n{'='*60}")
    print("STEP 1: LOADING EXCEL FILE")
    print(f"{'='*60}")
    print(f"📖 Loading: {excel_path}")
    
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    
    xls = pd.ExcelFile(excel_path)
    sheets = {}
    
    for sheet_name in xls.sheet_names:
        sheets[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
        print(f"   ✓ {sheet_name}: {len(sheets[sheet_name]):,} rows")
    
    return sheets


def merge_sheets_by_crd(
    sheets: Dict[str, pd.DataFrame],
    sheets_to_process: List[str],
    source_columns: Dict[str, str],
    crd_column: str = "organization_crd"
) -> pd.DataFrame:
    """Merge all specified sheets by CRD."""
    print(f"\n{'='*60}")
    print("STEP 2: MERGING SHEETS BY CRD")
    print(f"{'='*60}")
    
    crd_data = {}
    
    for sheet_name in sheets_to_process:
        if sheet_name not in sheets:
            continue
        
        df = sheets[sheet_name]
        source_col = source_columns.get(sheet_name)
        
        if source_col not in df.columns or crd_column not in df.columns:
            continue
        
        print(f"   📋 Processing {sheet_name}...")
        
        for _, row in df.iterrows():
            crd = row.get(crd_column)
            if pd.isna(crd):
                continue
            
            crd = str(crd).strip()
            if not crd:
                continue
            
            if crd not in crd_data:
                crd_data[crd] = {
                    "organization_crd": crd,
                    "company_name": row.get("company_name", ""),
                    "original_url": row.get("original_url", ""),
                    "services_raw": [],
                    "clients_raw": [],
                    "products_raw": []
                }
            
            value = row.get(source_col)
            if pd.notna(value) and str(value).strip():
                value_str = str(value).strip()
                
                if sheet_name == "Services":
                    crd_data[crd]["services_raw"].append(value_str)
                elif sheet_name == "Target_Clients":
                    crd_data[crd]["clients_raw"].append(value_str)
                elif sheet_name == "Products_Instruments":
                    crd_data[crd]["products_raw"].append(value_str)
    
    rows = []
    for crd, data in crd_data.items():
        rows.append({
            "organization_crd": data["organization_crd"],
            "company_name": data["company_name"],
            "original_url": data["original_url"],
            "services_raw": " | ".join(list(set(data["services_raw"]))),
            "clients_raw": " | ".join(list(set(data["clients_raw"]))),
            "products_raw": " | ".join(list(set(data["products_raw"])))
        })
    
    merged_df = pd.DataFrame(rows)
    print(f"   ✓ Merged {len(merged_df):,} unique CRDs")
    
    return merged_df


def create_llm_client(config: Dict) -> AzureOpenAI:
    """Create Azure OpenAI client."""
    print(f"\n{'='*60}")
    print("STEP 3: CREATING LLM CLIENT")
    print(f"{'='*60}")
    
    client = AzureOpenAI(
        azure_endpoint=config["azure_endpoint"],
        api_key=config["azure_api_key"],
        api_version=config["azure_api_version"]
    )
    
    print(f"   ✓ Client created for {config['azure_deployment']}")
    return client


# ============================================================================
# MAIN PROCESSING FUNCTION
# ============================================================================

def process_crds(
    df: pd.DataFrame,
    client: AzureOpenAI,
    config: Dict,
    limit: Optional[int] = None
) -> Dict[str, Dict]:
    """Process CRDs using 3-prompt approach."""
    
    print(f"\n{'='*60}")
    print("STEP 4: PROCESSING CRDs (3-PROMPT APPROACH)")
    print(f"{'='*60}")
    
    if limit:
        df = df.head(limit)
    
    all_results = {}
    total = len(df)
    
    for idx, row in df.iterrows():
        crd = str(row["organization_crd"])
        
        print(f"\n[{idx+1}/{total}] Processing CRD {crd}...")
        
        # Call all 3 prompts
        result = classify_crd(
            client=client,
            deployment=config["azure_deployment"],
            crd=crd,
            services_raw=row.get("services_raw", ""),
            clients_raw=row.get("clients_raw", ""),
            products_raw=row.get("products_raw", ""),
            max_retries=config["max_retries"]
        )
        
        # Validate
        validated, invalid = validate_classification(crd, result)
        
        if invalid:
            print(f"   ⚠️ Invalid values removed: {invalid[:3]}")
        
        all_results[crd] = validated
        
        # Show key results
        print(f"   ✓ service_category: {validated.get('service_category', '')[:50]}...")
        print(f"   ✓ client_category: {validated.get('client_category', '')[:50]}...")
        print(f"   ✓ asset_class: {validated.get('asset_class', '')[:50]}...")
        print(f"   ✓ geographic_exposure: {validated.get('geographic_exposure', '')}")
        print(f"   ✓ market_capitalization: {validated.get('market_capitalization', '')}")
        print(f"   ✓ tax_management_approach: {validated.get('tax_management_approach', '')}")
        print(f"   ✓ edge_cases: {validated.get('edge_cases', '')}")
        
        # Small delay between CRDs
        time.sleep(0.3)
    
    return all_results


# ============================================================================
# SAVE RESULTS
# ============================================================================
# ============================================================================
# SAVE RESULTS
# ============================================================================

def save_results(
    df: pd.DataFrame,
    results: Dict[str, Dict],
    output_path: str,
    mappings_path: str
):
    """Save results to CSV and JSON."""
    
    print(f"\n{'='*60}")
    print("STEP 5: SAVING RESULTS")
    print(f"{'='*60}")
    
    # Add classification columns to dataframe
    for col in OUTPUT_COLUMNS:
        df[col] = df["organization_crd"].astype(str).apply(
            lambda crd: results.get(crd, {}).get(col, "")
        )
    
    # Add additional_assets as JSON string
    df["additional_assets"] = df["organization_crd"].astype(str).apply(
        lambda crd: json.dumps(results.get(crd, {}).get("additional_assets", []))
    )
    
    # Merge additional_assets into main asset columns
    print("   📦 Merging additional_assets into main columns...")
    df = merge_additional_assets_into_columns(df)
    print("   ✓ Additional assets merged")
    
    # Save CSV
    csv_path = output_path.replace(".xlsx", ".csv")
    df.to_csv(csv_path, index=False)
    print(f"   ✓ Saved CSV: {csv_path}")
    
    # Save mappings JSON
    mappings_data = {
        "metadata": {
            "generated": datetime.now().isoformat(),
            "description": "RIA classification mappings (V7 - 3-prompt approach)",
            "total_mappings": len(results)
        },
        "mappings": results
    }
    
    with open(mappings_path, 'w', encoding='utf-8') as f:
        json.dump(mappings_data, f, indent=2, ensure_ascii=False)
    print(f"   ✓ Saved mappings: {mappings_path}")
    
    return df


# ============================================================================
# MAIN EXECUTION
# ============================================================================

if __name__ == "__main__":
    # Step 1: Load Excel
    sheets = load_excel_file(CONFIG["excel_path"])
    
    # Step 2: Merge by CRD
    merged_df = merge_sheets_by_crd(
        sheets,
        CONFIG["sheets_to_process"],
        CONFIG["source_columns"]
    )
    
    # Step 3: Create LLM client
    client = create_llm_client(CONFIG)
    
    # Step 4: Process CRDs (limit to 5 for testing)
    results = process_crds(merged_df, client, CONFIG)
    
    # Step 5: Save results
    final_df = save_results(
        merged_df,
        results,
        CONFIG["output_path"],
        CONFIG["mappings_file"]
    )
    
    # Show summary
    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")
    print(f"Total CRDs processed: {len(results)}")
    
    # Show sample result
    if results:
        sample_crd = list(results.keys())[0]
        print(f"\nSample result for CRD {sample_crd}:")
        for field, value in results[sample_crd].items():
            if value:
                print(f"   {field}: {value}")