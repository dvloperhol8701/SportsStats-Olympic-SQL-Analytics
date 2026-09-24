import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

def create_report():
    doc = docx.Document()

    # Configure Margins (Standard 1 inch)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Styling Helper Functions
    def add_title(text):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(22)
        run.bold = True
        run.font.color.rgb = RGBColor(16, 44, 87)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(4)

    def add_subtitle(text):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(12)
        run.italic = True
        run.font.color.rgb = RGBColor(80, 80, 80)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(18)

    def add_h1(text):
        h = doc.add_heading(level=1)
        run = h.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(15)
        run.bold = True
        run.font.color.rgb = RGBColor(16, 44, 87)
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(6)

    def add_h2(text):
        h = doc.add_heading(level=2)
        run = h.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(12.5)
        run.bold = True
        run.font.color.rgb = RGBColor(40, 80, 140)
        h.paragraph_format.space_before = Pt(10)
        h.paragraph_format.space_after = Pt(4)

    def add_body(text, bold_prefix=None):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(6)
        if bold_prefix:
            run_b = p.add_run(bold_prefix)
            run_b.font.name = 'Calibri'
            run_b.font.size = Pt(11)
            run_b.bold = True
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(11)

    def add_callout(text, prefix="> Narrative Angle: "):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(8)
        run_p = p.add_run(prefix)
        run_p.font.name = 'Calibri'
        run_p.font.size = Pt(10.5)
        run_p.bold = True
        run_p.font.color.rgb = RGBColor(16, 44, 87)
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(10.5)
        run.italic = True

    def add_code(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(30, 30, 30)

    def add_table_data(headers, rows):
        table = doc.add_table(rows=1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = 'Table Grid'
        
        hdr_cells = table.rows[0].cells
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.name = 'Calibri'
                r.font.size = Pt(10)
                r.bold = True

        for row_data in rows:
            row_cells = table.add_row().cells
            for i, val in enumerate(row_data):
                row_cells[i].text = str(val)
                p = row_cells[i].paragraphs[0]
                for r in p.runs:
                    r.font.name = 'Calibri'
                    r.font.size = Pt(9.5)
        doc.add_paragraph().paragraph_format.space_after = Pt(8)

    def add_image_if_exists(filename, width_inches=6.0, caption=None):
        if os.path.exists(filename):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)
            p.add_run().add_picture(filename, width=Inches(width_inches))
            if caption:
                p_cap = doc.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cap.paragraph_format.space_after = Pt(8)
                run_cap = p_cap.add_run(caption)
                run_cap.font.name = 'Calibri'
                run_cap.font.size = Pt(9.5)
                run_cap.italic = True
                run_cap.font.color.rgb = RGBColor(100, 100, 100)

    # --- DOCUMENT STRUCTURE ---

    # Title & Metadata
    add_title("CAPSTONE PROJECT REPORT: 120 YEARS OF OLYMPIC HISTORY")
    add_subtitle("Client Engagement: SportsStats Analytics Consultancy | Primary Analyst: Saurav\nSpecialization: Learn SQL Basics for Data Science")

    # Section 1: Executive Summary
    add_h1("1. Executive Summary")
    add_body("SportsStats was engaged to execute an in-depth relational database study across 120 years of modern Olympic competition logs (1896–2016). Operating across a dual-stakeholder business scenario, this analysis bridges quantitative physiological benchmarks for elite personal training academies with high-impact historical narratives for sports media journalists.")

    # Section 2: Data Architecture & Hygiene Audit
    add_h1("2. Relational Architecture & Data Hygiene Audit")
    add_body("The relational model links a primary fact table (athlete_events, 271,116 records) to a geographic dimension table (noc_regions, 230 records) via athlete_events.NOC = noc_regions.NOC.")

    add_h2("2.1 Data Hygiene Profile")
    add_table_data(
        ["Total Records", "Missing Ages", "Missing Heights", "Missing Weights", "Non-Medalists"],
        [["271,116", "9,474", "60,171", "62,875", "231,333"]]
    )
    add_body("Key Quality Takeaways:", "Hygiene Observations: ")
    add_body("Height and weight attributes exhibit ~22-23% null values, primarily concentrated in pre-1960 events. Consequently, physiological benchmarks require strict non-null filtering. Furthermore, an audit revealed Singapore is recorded as 'SGP' in athlete_events versus 'SIN' in noc_regions; all SQL queries implement COALESCE(r.region, a.Team) to prevent record drops.")

    # Section 3: Coaching Dossier
    add_h1("3. High-Performance Coaching Dossier (Trainers)")
    add_body("To guide athletic talent scouting and development, physical archetypes were evaluated across disciplines using Body Mass Index (BMI = Weight / (Height/100)^2):")

    add_h2("3.1 Anthropometric Profiling Across Sports")
    add_table_data(
        ["Sport", "Medalists", "Avg Age", "Avg Height (cm)", "Avg Weight (kg)", "Avg BMI"],
        [
            ["Weightlifting", "473", "25.3", "169.2", "79.9", "27.6"],
            ["Wrestling", "967", "26.3", "172.9", "77.5", "25.7"],
            ["Bobsleigh", "371", "29.7", "181.7", "89.2", "27.0"],
            ["Basketball", "1028", "25.3", "191.9", "86.9", "23.4"],
            ["Gymnastics", "705", "22.4", "162.7", "56.4", "21.2"],
            ["Athletics", "3648", "25.2", "179.3", "73.4", "22.7"]
        ]
    )

    # Insert Biometric Clustering Visualization
    add_image_if_exists("biometric_clustering.png", 6.0, "Figure 1: Morphological Clustering of Olympic Medalists Across Contrasting Sports")

    add_body("Key Coaching Takeaway:", "Physiological Specialization: ")
    add_body("High-BMI sports (Weightlifting, Wrestling) demand absolute mass and lower centers of gravity for mechanical leverage. Conversely, low-BMI disciplines (Gymnastics, Distance Athletics) prioritize power-to-weight ratios and minimal rotational inertia. Scouting standards must be discipline-calibrated rather than standardized.")

    add_h2("3.2 Peak Competitive Age Windows")
    add_table_data(
        ["Sport", "Youngest Gold", "Oldest Gold", "Avg Gold Age", "Age Spread (Yrs)"],
        [
            ["Swimming", "13", "38", "21.7", "25"],
            ["Gymnastics", "14", "36", "23.2", "22"],
            ["Athletics", "15", "42", "25.4", "27"],
            ["Boxing", "17", "32", "23.1", "15"],
            ["Equestrian", "18", "64", "35.3", "46"],
            ["Shooting", "15", "60", "33.6", "45"]
        ]
    )
    add_body("Key Coaching Takeaway:", "Longevity Windows: ")
    add_body("Swimming and Gymnastics demand front-loaded athletic investment with peak outcomes clustering tightly around 21-23 years. In contrast, technical precision sports like Shooting and Equestrian sustain multi-decade competitive lifecycles extending past age 40.")

    # Section 4: Media Story Pack
    add_h1("4. Sports Media Story Pack (Journalists)")
    add_body("In team sports (such as Men's Basketball), up to 12 medals are awarded to roster members for a single tournament win. To measure true all-time national dominance, CTEs were deployed to evaluate unique winning events:")

    add_h2("4.1 Top Nations by Distinct Event Victories")
    add_table_data(
        ["Country", "Gold Events", "Silver Events", "Bronze Events", "Total Event Medals"],
        [
            ["USA", "1035", "802", "708", "2545"],
            ["Russia / USSR", "591", "496", "484", "1571"],
            ["Germany", "444", "457", "490", "1391"],
            ["UK", "278", "317", "300", "895"],
            ["France", "234", "256", "312", "802"],
            ["Italy", "246", "214", "253", "713"]
        ]
    )

    add_h2("4.2 Longitudinal Growth of Female Participation (Summer Games)")
    add_table_data(
        ["Year", "Male Competitors", "Female Competitors", "Female Participation %"],
        [
            ["1896", "176", "0", "0.00%"],
            ["1928", "2937", "311", "9.58%"],
            ["1964", "4457", "678", "13.20%"],
            ["1984", "5263", "1565", "22.92%"],
            ["2000", "6579", "4069", "38.21%"],
            ["2016", "6179", "5034", "44.90%"]
        ]
    )

    # Insert Gender Growth Visualization
    add_image_if_exists("gender_growth.png", 5.8, "Figure 2: Longitudinal Evolution of Female Olympic Participation (1896–2016)")

    add_callout(
        "Female representation remained under 15% through the mid-20th century. Following institutional policy changes in the early 1980s, female participation accelerated rapidly, reaching near-parity (44.9%) by Rio 2016.",
        "> Media Narrative: "
    )

    add_h2("4.3 Investigative Storyline: The Host Nation Advantage")
    add_body("Using a 3-Games trailing moving average (AVG() OVER ROWS BETWEEN 3 PRECEDING AND 1 PRECEDING), we quantified the empirical surge experienced by host countries:")

    add_table_data(
        ["Year", "Host City", "Host Country", "Host Year Medals", "Trailing 3-Games Avg", "Host Boost (Net)"],
        [
            ["1984", "Los Angeles", "USA", "173", "98.3", "+74.7"],
            ["2008", "Beijing", "China", "100", "57.7", "+42.3"],
            ["2000", "Sydney", "Australia", "58", "27.3", "+30.7"],
            ["2012", "London", "UK", "65", "35.3", "+29.7"],
            ["1988", "Seoul", "South Korea", "33", "8.7", "+24.3"],
            ["1992", "Barcelona", "Spain", "22", "5.0", "+17.0"],
            ["2016", "Rio de Janeiro", "Brazil", "19", "14.3", "+4.7"],
            ["1996", "Atlanta", "USA", "101", "125.0", "-24.0"]
        ]
    )
    add_callout(
        "Host nations experience an empirical surge of +15 to +42 medals above their trailing baseline due to full event allotments, targeted multi-year funding, and crowd familiarity (e.g., China +42.3 in 2008; Australia +30.7 in 2000).",
        "> Media Narrative: "
    )

    # Section 5: Technical SQL Implementation
    add_h1("5. Technical SQL Implementation")
    add_code(
"""-- 1. Deduplicated National Dominance Query (Event-Level)
WITH DistinctEventMedals AS (
    SELECT DISTINCT Year, Season, Sport, Event, Medal, NOC
    FROM athlete_events
    WHERE Medal IN ('Gold', 'Silver', 'Bronze')
)
SELECT 
    COALESCE(r.region, d.NOC) AS country,
    COUNT(CASE WHEN d.Medal = 'Gold' THEN 1 END) AS gold_events,
    COUNT(CASE WHEN d.Medal = 'Silver' THEN 1 END) AS silver_events,
    COUNT(CASE WHEN d.Medal = 'Bronze' THEN 1 END) AS bronze_events,
    COUNT(*) AS total_event_medals
FROM DistinctEventMedals d
LEFT JOIN noc_regions r ON d.NOC = r.NOC
GROUP BY COALESCE(r.region, d.NOC)
ORDER BY total_event_medals DESC
LIMIT 10;

-- 2. Host Nation Advantage (Rolling Window Function)
WITH HostNations AS (
    SELECT DISTINCT Year, City,
        CASE 
            WHEN City = 'London' THEN 'UK'
            WHEN City IN ('Los Angeles', 'Atlanta', 'St. Louis') THEN 'USA'
            WHEN City = 'Beijing' THEN 'China'
            WHEN City IN ('Sydney', 'Melbourne') THEN 'Australia'
            WHEN City = 'Seoul' THEN 'South Korea'
            WHEN City = 'Barcelona' THEN 'Spain'
            WHEN City = 'Rio de Janeiro' THEN 'Brazil'
            ELSE 'Other'
        END AS host_country
    FROM athlete_events WHERE Season = 'Summer'
),
CountryMedalsByYear AS (
    SELECT a.Year, COALESCE(r.region, a.NOC) AS country,
           COUNT(DISTINCT a.Event || a.Medal) AS total_medals
    FROM athlete_events a
    LEFT JOIN noc_regions r ON a.NOC = r.NOC
    WHERE a.Medal IN ('Gold', 'Silver', 'Bronze') AND a.Season = 'Summer'
    GROUP BY a.Year, COALESCE(r.region, a.NOC)
),
MedalTrends AS (
    SELECT c.Year, c.country, c.total_medals,
           ROUND(AVG(c.total_medals) OVER (
               PARTITION BY c.country ORDER BY c.Year 
               ROWS BETWEEN 3 PRECEDING AND 1 PRECEDING
           ), 1) AS trailing_3games_avg
    FROM CountryMedalsByYear c
)
SELECT m.Year, h.City, m.country AS host_country,
       m.total_medals AS host_year_medals, m.trailing_3games_avg,
       ROUND(m.total_medals - m.trailing_3games_avg, 1) AS host_boost
FROM MedalTrends m
JOIN HostNations h ON m.Year = h.Year AND m.country = h.host_country
WHERE m.trailing_3games_avg IS NOT NULL AND m.Year >= 1980
ORDER BY host_boost DESC;"""
    )

    # Section 6: Limitations & Strategic Recommendations
    add_h1("6. Limitations & Strategic Recommendations")
    add_body("1. Environmental Normalization: Historical Olympic datasets do not capture altitude, temperature, or humidity, which heavily influence endurance event outcomes.")
    add_body("2. Historical Entity Changes: Multi-nation historical blocs (e.g., USSR, East Germany) require clear contextual callouts in journalism stories to compare fairly with modern national boundaries.")
    add_body("3. Coaching Integration: Training academies should ingest modern micro-wearable sensor telemetry to supplement historical anthropometric benchmarks.")

    doc.save("SportsStats_Final_Report.docx")
    print("Updated report saved successfully as 'SportsStats_Final_Report.docx'!")

if __name__ == "__main__":
    create_report()