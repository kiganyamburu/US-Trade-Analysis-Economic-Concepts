"""
US Trade Analysis & Economic Concepts - Excel Workbook Generator
Creates an Excel workbook with 5 worksheets analyzing US trade data.

Data Sources:
- US Census Bureau Foreign Trade Division (https://www.census.gov/foreign-trade/)
- Bureau of Economic Analysis (https://www.bea.gov/data/intl-trade-investment/international-trade-goods-and-services)
- Office of the United States Trade Representative (https://ustr.gov/countries-regions)
"""

import xlsxwriter
from datetime import datetime

# Create workbook
workbook = xlsxwriter.Workbook("US_Trade_Analysis_2024.xlsx")

# ============================================================================
# DEFINE FORMATS
# ============================================================================
title_format = workbook.add_format(
    {
        "bold": True,
        "font_size": 16,
        "align": "center",
        "valign": "vcenter",
        "font_color": "white",
        "bg_color": "#1F4E79",
        "border": 1,
    }
)
header_format = workbook.add_format(
    {
        "bold": True,
        "font_size": 11,
        "align": "center",
        "valign": "vcenter",
        "bg_color": "#D6DCE4",
        "border": 1,
        "text_wrap": True,
    }
)
data_format = workbook.add_format(
    {"font_size": 11, "align": "center", "valign": "vcenter", "border": 1}
)
currency_format = workbook.add_format(
    {
        "font_size": 11,
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "num_format": "$#,##0.0",
    }
)
percent_format = workbook.add_format(
    {
        "font_size": 11,
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "num_format": "0.0%",
    }
)
source_format = workbook.add_format(
    {"font_size": 9, "italic": True, "align": "left", "valign": "vcenter"}
)
question_format = workbook.add_format(
    {
        "bold": True,
        "font_size": 12,
        "text_wrap": True,
        "valign": "top",
        "bg_color": "#E2EFDA",
        "border": 2,
    }
)
answer_format = workbook.add_format(
    {"font_size": 11, "text_wrap": True, "valign": "top", "border": 1}
)

# ============================================================================
# WORKSHEET 1: US IMPORTS BY INDUSTRY SECTOR (2024 Data)
# ============================================================================
ws1 = workbook.add_worksheet("Imports by Industry")
ws1.set_column("A:A", 35)
ws1.set_column("B:C", 18)

# Data: US Imports by End-Use Category 2024 (in billions USD)
# Source: US Census Bureau FT-900, October 2025 Release
imports_by_sector = [
    ("Capital Goods (exc. automotive)", 753.2, 0.232),
    ("Consumer Goods", 745.8, 0.230),
    ("Industrial Supplies & Materials", 698.4, 0.215),
    ("Automotive Vehicles & Parts", 469.1, 0.145),
    ("Foods, Feeds & Beverages", 198.6, 0.061),
    ("Other Goods", 378.9, 0.117),
]

# Title
ws1.merge_range(
    "A1:C1",
    "Figure 1: Composition of US Imports by Industry Sector, 2024",
    title_format,
)
ws1.set_row(0, 30)

# Headers
ws1.write("A3", "Industry Sector", header_format)
ws1.write("B3", "Value (Billions USD)", header_format)
ws1.write("C3", "Percentage of Total", header_format)

# Data
for i, (sector, value, pct) in enumerate(imports_by_sector, start=3):
    ws1.write(i, 0, sector, data_format)
    ws1.write(i, 1, value, currency_format)
    ws1.write(i, 2, pct, percent_format)

# Total
total_imports = sum([x[1] for x in imports_by_sector])
ws1.write(9, 0, "TOTAL", header_format)
ws1.write(9, 1, total_imports, currency_format)
ws1.write(9, 2, 1.0, percent_format)

# Source citation
ws1.write(
    "A12",
    "Source: U.S. Census Bureau, Foreign Trade Division, FT-900 Report (January 2025)",
    source_format,
)
ws1.write(
    "A13",
    "Data URL: https://www.census.gov/foreign-trade/Press-Release/current_press_release/index.html",
    source_format,
)

# Create pie chart
chart1 = workbook.add_chart({"type": "pie"})
chart1.add_series(
    {
        "name": "US Imports by Industry Sector",
        "categories": "='Imports by Industry'!$A$4:$A$9",
        "values": "='Imports by Industry'!$B$4:$B$9",
        "data_labels": {"percentage": True, "category": False, "font": {"size": 9}},
    }
)
chart1.set_title(
    {
        "name": "Figure 1: US Imports by Industry Sector, 2024\n(Billions of USD)",
        "name_font": {"size": 12, "bold": True},
    }
)
chart1.set_legend({"position": "right", "font": {"size": 9}})
chart1.set_size({"width": 550, "height": 400})
ws1.insert_chart("E3", chart1)

# ============================================================================
# WORKSHEET 2: US EXPORTS BY INDUSTRY SECTOR (2024 Data)
# ============================================================================
ws2 = workbook.add_worksheet("Exports by Industry")
ws2.set_column("A:A", 35)
ws2.set_column("B:C", 18)

# Data: US Exports by End-Use Category 2024 (in billions USD)
# Source: US Census Bureau FT-900, October 2025 Release
exports_by_sector = [
    ("Capital Goods (exc. automotive)", 586.7, 0.286),
    ("Industrial Supplies & Materials", 574.3, 0.280),
    ("Consumer Goods", 253.8, 0.124),
    ("Automotive Vehicles & Parts", 186.2, 0.091),
    ("Foods, Feeds & Beverages", 192.5, 0.094),
    ("Other Goods", 256.5, 0.125),
]

# Title
ws2.merge_range(
    "A1:C1",
    "Figure 2: Composition of US Exports by Industry Sector, 2024",
    title_format,
)
ws2.set_row(0, 30)

# Headers
ws2.write("A3", "Industry Sector", header_format)
ws2.write("B3", "Value (Billions USD)", header_format)
ws2.write("C3", "Percentage of Total", header_format)

# Data
for i, (sector, value, pct) in enumerate(exports_by_sector, start=3):
    ws2.write(i, 0, sector, data_format)
    ws2.write(i, 1, value, currency_format)
    ws2.write(i, 2, pct, percent_format)

# Total
total_exports = sum([x[1] for x in exports_by_sector])
ws2.write(9, 0, "TOTAL", header_format)
ws2.write(9, 1, total_exports, currency_format)
ws2.write(9, 2, 1.0, percent_format)

# Source citation
ws2.write(
    "A12",
    "Source: U.S. Census Bureau, Foreign Trade Division, FT-900 Report (January 2025)",
    source_format,
)
ws2.write(
    "A13",
    "Data URL: https://www.census.gov/foreign-trade/Press-Release/current_press_release/index.html",
    source_format,
)

# Create pie chart
chart2 = workbook.add_chart({"type": "pie"})
chart2.add_series(
    {
        "name": "US Exports by Industry Sector",
        "categories": "='Exports by Industry'!$A$4:$A$9",
        "values": "='Exports by Industry'!$B$4:$B$9",
        "data_labels": {"percentage": True, "category": False, "font": {"size": 9}},
    }
)
chart2.set_title(
    {
        "name": "Figure 2: US Exports by Industry Sector, 2024\n(Billions of USD)",
        "name_font": {"size": 12, "bold": True},
    }
)
chart2.set_legend({"position": "right", "font": {"size": 9}})
chart2.set_size({"width": 550, "height": 400})
ws2.insert_chart("E3", chart2)

# ============================================================================
# WORKSHEET 3: US IMPORTS BY TRADING PARTNER (Top 5, 2024 Data)
# ============================================================================
ws3 = workbook.add_worksheet("Imports by Partner")
ws3.set_column("A:A", 25)
ws3.set_column("B:C", 18)

# Data: US Imports by Top 5 Trading Partners 2024 (in billions USD)
# Source: US Census Bureau Foreign Trade - Top Trading Partners
imports_by_partner = [
    ("China", 427.2, 0.142),
    ("Mexico", 505.8, 0.168),
    ("Canada", 412.3, 0.137),
    ("Japan", 135.2, 0.045),
    ("Germany", 157.6, 0.052),
    ("All Other Countries", 1373.9, 0.456),
]

# Title
ws3.merge_range(
    "A1:C1", "Figure 3: US Imports by Top 5 Trading Partners, 2024", title_format
)
ws3.set_row(0, 30)

# Headers
ws3.write("A3", "Trading Partner", header_format)
ws3.write("B3", "Value (Billions USD)", header_format)
ws3.write("C3", "Percentage of Total", header_format)

# Data
for i, (partner, value, pct) in enumerate(imports_by_partner, start=3):
    ws3.write(i, 0, partner, data_format)
    ws3.write(i, 1, value, currency_format)
    ws3.write(i, 2, pct, percent_format)

# Total
total_imports_partner = sum([x[1] for x in imports_by_partner])
ws3.write(10, 0, "TOTAL", header_format)
ws3.write(10, 1, total_imports_partner, currency_format)
ws3.write(10, 2, 1.0, percent_format)

# Source citation
ws3.write(
    "A13",
    "Source: U.S. Census Bureau, Foreign Trade Division - Top Trading Partners (2024)",
    source_format,
)
ws3.write(
    "A14",
    "Data URL: https://www.census.gov/foreign-trade/statistics/highlights/toppartners.html",
    source_format,
)

# Create pie chart
chart3 = workbook.add_chart({"type": "pie"})
chart3.add_series(
    {
        "name": "US Imports by Trading Partner",
        "categories": "='Imports by Partner'!$A$4:$A$9",
        "values": "='Imports by Partner'!$B$4:$B$9",
        "data_labels": {"percentage": True, "category": False, "font": {"size": 9}},
    }
)
chart3.set_title(
    {
        "name": "Figure 3: US Imports by Top 5 Trading Partners, 2024\n(Billions of USD)",
        "name_font": {"size": 12, "bold": True},
    }
)
chart3.set_legend({"position": "right", "font": {"size": 9}})
chart3.set_size({"width": 550, "height": 400})
ws3.insert_chart("E3", chart3)

# ============================================================================
# WORKSHEET 4: US EXPORTS BY TRADING PARTNER (Top 5, 2024 Data)
# ============================================================================
ws4 = workbook.add_worksheet("Exports by Partner")
ws4.set_column("A:A", 25)
ws4.set_column("B:C", 18)

# Data: US Exports by Top 5 Trading Partners 2024 (in billions USD)
# Source: US Census Bureau Foreign Trade - Top Trading Partners
exports_by_partner = [
    ("Canada", 351.8, 0.171),
    ("Mexico", 322.5, 0.157),
    ("China", 143.5, 0.070),
    ("Japan", 79.8, 0.039),
    ("United Kingdom", 76.4, 0.037),
    ("All Other Countries", 1076.0, 0.526),
]

# Title
ws4.merge_range(
    "A1:C1", "Figure 4: US Exports by Top 5 Trading Partners, 2024", title_format
)
ws4.set_row(0, 30)

# Headers
ws4.write("A3", "Trading Partner", header_format)
ws4.write("B3", "Value (Billions USD)", header_format)
ws4.write("C3", "Percentage of Total", header_format)

# Data
for i, (partner, value, pct) in enumerate(exports_by_partner, start=3):
    ws4.write(i, 0, partner, data_format)
    ws4.write(i, 1, value, currency_format)
    ws4.write(i, 2, pct, percent_format)

# Total
total_exports_partner = sum([x[1] for x in exports_by_partner])
ws4.write(10, 0, "TOTAL", header_format)
ws4.write(10, 1, total_exports_partner, currency_format)
ws4.write(10, 2, 1.0, percent_format)

# Source citation
ws4.write(
    "A13",
    "Source: U.S. Census Bureau, Foreign Trade Division - Top Trading Partners (2024)",
    source_format,
)
ws4.write(
    "A14",
    "Data URL: https://www.census.gov/foreign-trade/statistics/highlights/toppartners.html",
    source_format,
)

# Create pie chart
chart4 = workbook.add_chart({"type": "pie"})
chart4.add_series(
    {
        "name": "US Exports by Trading Partner",
        "categories": "='Exports by Partner'!$A$4:$A$9",
        "values": "='Exports by Partner'!$B$4:$B$9",
        "data_labels": {"percentage": True, "category": False, "font": {"size": 9}},
    }
)
chart4.set_title(
    {
        "name": "Figure 4: US Exports by Top 5 Trading Partners, 2024\n(Billions of USD)",
        "name_font": {"size": 12, "bold": True},
    }
)
chart4.set_legend({"position": "right", "font": {"size": 9}})
chart4.set_size({"width": 550, "height": 400})
ws4.insert_chart("E3", chart4)

# ============================================================================
# WORKSHEET 5: ECONOMIC CONCEPTS - SHORT WRITTEN RESPONSES
# ============================================================================
ws5 = workbook.add_worksheet("Economic Questions")
ws5.set_column("A:A", 100)
ws5.set_row(0, 25)

# Title
ws5.merge_range(
    "A1:A1",
    "Part 2: Short Written Responses - International Trade Theory & Policy",
    title_format,
)

# Questions and Answers
questions_answers = [
    # Question 1
    (
        "QUESTION 1: Historical Trends\nHow has the fundamental nature of United States trade evolved over the past 100 years? Discuss shifts in volume, composition, or partners.",
        """ANSWER:

Over the past century, U.S. trade has undergone profound transformations across three key dimensions:

VOLUME: Trade has expanded dramatically from approximately $9 billion in total trade in 1920 to over $7 trillion today. This 700-fold increase reflects globalization, reduced transportation costs, trade liberalization (GATT/WTO), and technological advances in logistics and communication.

COMPOSITION:
• 1920s-1950s: The U.S. was primarily an exporter of agricultural products and raw materials, while importing manufactured goods.
• 1960s-1980s: A shift toward exporting manufactured goods (automobiles, machinery, electronics) as the U.S. became an industrial powerhouse.
• 1990s-Present: The economy has transitioned to services and high-technology exports (software, intellectual property, financial services, aerospace). Meanwhile, imports have shifted to consumer goods and electronics from Asia.

TRADING PARTNERS:
• Early 20th Century: Trade was concentrated with European nations (UK, Germany, France).
• Mid-20th Century: Japan emerged as a major partner following WWII reconstruction.
• Late 20th Century-Present: China became the largest source of imports; NAFTA/USMCA made Canada and Mexico top partners. Today's top 5 partners (Mexico, Canada, China, Japan, Germany) account for over 50% of all U.S. trade.""",
    ),
    # Question 2
    (
        "QUESTION 2: Opportunity Cost\nDefine the concept of 'opportunity cost' specifically as it applies to a country's industrial policy. Provide concrete examples to illustrate your definition.",
        """ANSWER:

DEFINITION: Opportunity cost in industrial policy refers to the value of the next-best alternative foregone when a nation chooses to allocate its limited resources (capital, labor, land, technology) toward one industry or sector instead of another.

APPLICATION TO INDUSTRIAL POLICY:
When a government decides to promote or protect a specific industry, it diverts resources that could have been used elsewhere. The true cost is not just the direct expenditure but what that investment could have produced in alternative uses.

CONCRETE EXAMPLES:

1. Agricultural Subsidies vs. Technology Investment
   • When the U.S. allocates $20 billion annually to farm subsidies, the opportunity cost includes potential investments in semiconductor manufacturing, clean energy R&D, or workforce training programs that might yield higher long-term economic returns.

2. Steel Tariffs (Section 232, 2018)
   • Protecting domestic steel through tariffs preserved approximately 8,700 steel jobs.
   • Opportunity cost: Higher steel prices increased costs for downstream industries (automobiles, construction, appliances), which employ 80x more workers than steel production. Studies estimated 75,000+ jobs were lost in steel-consuming industries.

3. Defense Industrial Base vs. Infrastructure
   • Resources allocated to defense manufacturing represent opportunity costs in civilian infrastructure, healthcare, or education investments that could improve long-term productivity.

4. China's Industrial Policy
   • China's massive investment in solar panel manufacturing meant reduced investment in other sectors like biotechnology or consumer services, but positioned them as the global leader in renewable energy equipment.""",
    ),
    # Question 3
    (
        "QUESTION 3: Advantage Types\nDistinguish between a country's absolute advantage and comparative advantage. Provide examples of each to demonstrate the difference.",
        """ANSWER:

ABSOLUTE ADVANTAGE (Adam Smith):
A country has an absolute advantage when it can produce a good using fewer resources (labor, capital, time) than another country. It's about being "better" in absolute terms.

COMPARATIVE ADVANTAGE (David Ricardo):
A country has a comparative advantage when it can produce a good at a lower opportunity cost than another country. Even if a country is worse at producing everything, it still benefits from specializing in what it's "least bad" at.

KEY DISTINCTION:
• Absolute advantage compares productivity levels
• Comparative advantage compares opportunity costs

EXAMPLES:

ABSOLUTE ADVANTAGE:
• Saudi Arabia has an absolute advantage in oil production – it can extract oil at $10-15 per barrel, while U.S. shale oil costs $35-50 per barrel.
• The United States has an absolute advantage in commercial aircraft manufacturing (Boeing) due to advanced technology, skilled engineers, and established supply chains.

COMPARATIVE ADVANTAGE:
Consider two countries producing wine and cloth:

                    Wine (hours/unit)    Cloth (hours/unit)
   Portugal              80                   90
   England              120                  100

• Portugal has absolute advantage in BOTH goods (fewer hours for each)
• But comparative advantage differs:
   - Portugal's opportunity cost of wine = 80/90 = 0.89 cloth
   - England's opportunity cost of wine = 120/100 = 1.20 cloth
   - Portugal has comparative advantage in wine (lower opportunity cost)
   - England has comparative advantage in cloth (lower opportunity cost)

REAL-WORLD EXAMPLE:
The U.S. could produce textiles domestically, but the opportunity cost is high (those workers could be producing semiconductors, software, or pharmaceuticals). Bangladesh has comparative advantage in textiles because its opportunity cost of textile production is lower – its alternative employment opportunities in high-tech sectors are limited.""",
    ),
    # Question 4
    (
        "QUESTION 4: Labor and Trade Policy\na) Do low-wage countries possess an 'unfair' advantage in the production of goods?\nb) Should this wage disparity influence US trade policy?\nc) Explain how labor productivity should theoretically drive wage levels.",
        """ANSWER:

a) DO LOW-WAGE COUNTRIES POSSESS AN "UNFAIR" ADVANTAGE?

No, low wages do not constitute an "unfair" advantage – they reflect legitimate differences in:
• Labor productivity levels
• Cost of living and purchasing power parity
• Stage of economic development
• Factor endowments (abundant labor vs. scarce capital)

Low wages are often offset by:
• Lower productivity (output per worker-hour)
• Infrastructure limitations
• Higher logistics/transportation costs
• Quality control challenges
• Intellectual property risks

EXAMPLE: A U.S. worker earning $25/hour may produce 10 units; a Vietnamese worker earning $2.50/hour may produce 1 unit. The effective labor cost per unit is identical.

b) SHOULD WAGE DISPARITY INFLUENCE U.S. TRADE POLICY?

Trade policy should NOT be based primarily on wage differences because:
• Wages reflect productivity, not unfairness
• Protectionism harms U.S. consumers through higher prices
• Retaliation damages U.S. export industries
• It violates comparative advantage principles

HOWEVER, trade policy SHOULD address:
• Currency manipulation (artificial wage suppression)
• Labor rights violations (forced labor, unsafe conditions)
• Environmental standards circumvention
• Intellectual property theft
• Government subsidies that distort competition

c) HOW LABOR PRODUCTIVITY DRIVES WAGES (THEORETICAL FRAMEWORK):

Marginal Productivity Theory: In competitive markets, wages tend to equal the marginal revenue product of labor (MRPL) – the additional revenue a firm earns from hiring one more worker.

Wage = Marginal Physical Product × Price of Output

• High productivity → Higher wages (workers generate more value)
• Low productivity → Lower wages (workers generate less value)

EMPIRICAL EVIDENCE:
• U.S. manufacturing productivity rose 3.4% annually (1950-2000); real wages tracked closely
• Countries that invest in education, technology, and infrastructure see wage increases
• "Middle-income trap" occurs when productivity stagnates despite rising wages""",
    ),
    # Question 5
    (
        "QUESTION 5: Specialization\nUnder what circumstances, if ever, should a country focus on complete specialization (producing a strictly restricted range of products) to enhance efficiency?",
        """ANSWER:

ARGUMENTS FOR COMPLETE SPECIALIZATION:

1. Small Economies with Limited Resources
   • Small nations (Singapore, Luxembourg) cannot achieve economies of scale across multiple industries
   • Focusing on niche sectors maximizes efficiency
   • Example: Singapore specializes in finance, shipping, and petrochemicals

2. Extreme Comparative Advantage
   • When opportunity costs are vastly different, specialization maximizes global welfare
   • Example: Saudi Arabia's near-total specialization in oil/petrochemicals

3. Increasing Returns to Scale
   • Industries with significant learning curves benefit from concentrated production
   • Example: Taiwan's semiconductor foundries (TSMC)

ARGUMENTS AGAINST COMPLETE SPECIALIZATION:

1. Volatility and Risk
   • Commodity-dependent economies suffer from price swings
   • Example: Venezuela's oil dependence led to economic collapse when prices fell

2. Terms of Trade Deterioration
   • Primary product exporters face declining relative prices over time (Prebisch-Singer hypothesis)

3. National Security Concerns
   • Complete dependence on imports for essential goods (food, energy, defense) creates vulnerability
   • COVID-19 revealed risks of over-reliance on foreign pharmaceutical and medical supply chains

4. Employment and Social Stability
   • Sudden shifts can cause structural unemployment
   • Diversification provides economic resilience

5. Dutch Disease
   • Booming export sectors can appreciate currency, harming other industries

OPTIMAL APPROACH: MANAGED DIVERSIFICATION

Countries should:
• Specialize in areas of strong comparative advantage
• Maintain strategic domestic capacity in essential sectors
• Diversify to manage risk while capturing efficiency gains
• Use trade policy to facilitate gradual transitions, not permanent protection

EXAMPLE: South Korea transitioned from textiles → shipbuilding → electronics → semiconductors → content/entertainment, maintaining specialization within each phase while building new comparative advantages.""",
    ),
    # Question 6
    (
        "QUESTION 6: Efficiency Types\nDefine and differentiate between allocative efficiency and productive efficiency. Provide specific examples for each concept.",
        """ANSWER:

PRODUCTIVE EFFICIENCY:
Producing goods and services at the lowest possible cost, using the optimal combination of inputs. This occurs when a firm/economy operates on its production possibilities frontier (PPF).

FORMULA: Minimize Cost = f(Labor, Capital, Materials) for given output

CHARACTERISTICS:
• No waste of resources
• Using best available technology
• Operating at minimum average total cost
• Cannot produce more of one good without sacrificing another

EXAMPLES OF PRODUCTIVE EFFICIENCY:
1. Toyota Production System: Lean manufacturing minimizes inventory, defects, and idle time – producing vehicles at lower cost than competitors
2. Amazon Fulfillment: Automated warehouses with robotics and AI-optimized logistics achieve lower per-package costs
3. Agricultural Modernization: U.S. farms produce 3x more output with 1/3 the labor compared to 1950

ALLOCATIVE EFFICIENCY:
Producing the combination of goods and services that maximizes social welfare by aligning production with consumer preferences. Resources flow to their highest-valued uses.

FORMULA: Price = Marginal Cost (P = MC) for all goods

CHARACTERISTICS:
• Resources directed to most valued uses
• Consumer surplus + producer surplus maximized
• Market prices reflect true social costs and benefits
• Optimal mix of goods produced

EXAMPLES OF ALLOCATIVE EFFICIENCY:
1. Price Signals in Energy Markets: When oil prices rise, consumers reduce consumption and producers increase exploration – resources reallocate efficiently
2. Carbon Pricing: A carbon tax internalizes pollution costs, shifting resources from high-emission to low-emission production
3. Free Trade: Removing tariffs allows resources to flow to countries with comparative advantage, increasing global welfare

KEY DIFFERENCES:

| Aspect              | Productive Efficiency        | Allocative Efficiency          |
|---------------------|------------------------------|--------------------------------|
| Focus               | HOW to produce               | WHAT to produce                |
| Question            | Are we minimizing costs?     | Are we producing the right mix?|
| Graphical           | On the PPF curve             | At the right point on the PPF  |
| Market Condition    | Firms minimize costs         | P = MC in all markets          |
| Failure Example     | Inefficient factories        | Monopoly (P > MC)              |

IMPORTANT RELATIONSHIP:
An economy can be productively efficient but allocatively inefficient. Example: The Soviet Union efficiently produced millions of tanks (productive efficiency) but consumers wanted cars and consumer goods (allocative inefficiency).""",
    ),
]

row = 3
for q_num, (question, answer) in enumerate(questions_answers, start=1):
    # Question box
    ws5.set_row(row - 1, 60)  # Height for question
    ws5.write(row - 1, 0, question, question_format)

    # Answer box
    answer_height = max(200, len(answer) // 3)  # Approximate row height
    ws5.set_row(row, answer_height)
    ws5.write(row, 0, answer, answer_format)

    row += 3  # Skip a row between Q&A pairs

# Data Sources Summary
ws5.write(row, 0, "DATA SOURCES FOR THIS WORKBOOK:", header_format)
ws5.set_row(row, 20)
row += 1

sources = [
    "1. U.S. Census Bureau, Foreign Trade Division: https://www.census.gov/foreign-trade/",
    "2. Bureau of Economic Analysis, International Trade in Goods and Services: https://www.bea.gov/data/intl-trade-investment/international-trade-goods-and-services",
    "3. Office of the United States Trade Representative: https://ustr.gov/countries-regions",
    "4. Federal Reserve Economic Data (FRED): https://fred.stlouisfed.org/",
    "",
    "Report Generated: " + datetime.now().strftime("%B %d, %Y"),
]

for source in sources:
    ws5.write(row, 0, source, source_format)
    row += 1

# Close workbook
workbook.close()

print("✅ Excel workbook 'US_Trade_Analysis_2024.xlsx' created successfully!")
print("\nWorkbook contains 5 worksheets:")
print("  1. Imports by Industry - US imports by industrial sector with pie chart")
print("  2. Exports by Industry - US exports by industrial sector with pie chart")
print("  3. Imports by Partner - US imports by top 5 trading partners with pie chart")
print("  4. Exports by Partner - US exports by top 5 trading partners with pie chart")
print("  5. Economic Questions - 6 economic theory questions with detailed answers")
print("\nData Sources:")
print("  • U.S. Census Bureau Foreign Trade Division")
print("  • Bureau of Economic Analysis")
print("  • Office of the United States Trade Representative")
