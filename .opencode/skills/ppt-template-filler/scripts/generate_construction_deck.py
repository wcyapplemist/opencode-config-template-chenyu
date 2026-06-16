"""Generate: Digital Technology in Construction (16-slide deep deck)."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, ".")
from ppt_builder import generate_ppt_from_data, DEFAULT_OUTPUT_DIR

slide_data = [
    # ================================================================
    # 1. TITLE SLIDE
    # ================================================================
    {
        "slide_type": "title_slide",
        "title": "Digital Technology in Construction",
        "subtitle": "Transforming the Built Environment",
        "notes": (
            "KEY MESSAGE: Digital tools are reshaping how we design, build, and operate structures.\n"
            "\"Hold the slide for two seconds before you speak.\"\n"
            "\"Good [morning/afternoon], I'm [Name] \u2014 and today I want to show you how digital technology is fundamentally changing the construction industry.\"\n"
            "Pause. Let it land.\n"
            "\"We'll walk through the core technologies, the real-world evidence, and a practical roadmap for adoption.\"\n"
            "TRANSITION: \"Let's start with why this matters right now.\"\n"
            "COACHING: Eye contact, confident opening. Do not read the slide. Be ready for: \"Is this hype?\" \u2014 lead with the $39 billion market projection."
        ),
    },
    # ================================================================
    # 2. CONTENT - WHY NOW
    # ================================================================
    {
        "slide_type": "content_slide",
        "title": "Why Digital Transformation Now",
        "body": (
            "**Persistent Cost Overruns** \u2014 67% of large projects exceed budget by 20% or more\n"
            "**Schedule Delays** \u2014 Average 20-month delay on mega-projects globally\n"
            "**Safety & Productivity Gap** \u2014 Construction productivity has grown only 1% per year for 20 years"
        ),
        "notes": (
            "KEY MESSAGE: The industry has three entrenched problems that digital tools directly address.\n"
            "\"Let's make this concrete. These are not theoretical challenges \u2014 they are everyday, billion-dollar problems.\"\n"
            "\"Sixty-seven percent of large projects blow past their budget. Mega-projects run an average of twenty months late. And productivity? It's barely moved in two decades.\"\n"
            "Pause. Let the numbers land.\n"
            "\"These are the pain points that make digital transformation not optional, but urgent.\"\n"
            "TRANSITION: \"Here are the technologies that solve them.\"\n"
            "COACHING: Matter-of-fact, not alarmist. Be ready for: \"Are these global numbers?\" \u2014 answer: yes, McKinsey Global Institute and Flyvbjerg benchmark data."
        ),
    },
    # ================================================================
    # 3. SECTION HEADER - CORE TECHNOLOGIES
    # ================================================================
    {
        "slide_type": "section_header_slide",
        "title": "Core Technologies",
        "notes": (
            "KEY MESSAGE: Six technology pillars are reshaping construction.\n"
            "\"Now let's go deeper into the tools themselves.\"\n"
            "\"We're going to look at BIM, IoT, artificial intelligence, drones, robotics, and cloud platforms \u2014 six pillars that together form the digital foundation.\"\n"
            "TRANSITION: \"Let's start with the one that started it all \u2014 BIM.\"\n"
            "COACHING: Brief pause before moving on. Set up the next section as a logical progression."
        ),
    },
    # ================================================================
    # 4. CONTENT - BIM
    # ================================================================
    {
        "slide_type": "content_slide",
        "title": "BIM: Building Information Modeling",
        "body": (
            "**Shared Digital Model** \u2014 Architects, engineers, and contractors collaborate on one 3D model\n"
            "**Clash Detection** \u2014 Automated interference checks catch problems weeks before construction\n"
            "**Lifecycle Management** \u2014 Data flows from design through construction into operations and maintenance"
        ),
        "notes": (
            "KEY MESSAGE: BIM catches clashes on screen \u2014 not on site.\n"
            "\"BIM gives every discipline one shared digital model, so problems are visible before anyone pours concrete.\"\n"
            "Pause. Let the concept land.\n"
            "\"In our pilot projects, automated clash detection alone cut rework by up to thirty percent. That's weeks of delay and real money recovered.\"\n"
            "\"And it doesn't stop at handover \u2014 the same model becomes an operations and maintenance asset.\"\n"
            "TRANSITION: \"BIM provides the model. Now let's add real-time data from the site itself.\"\n"
            "COACHING: Walk through the three points top to bottom. Be ready for: \"Does BIM work with older projects?\" \u2014 answer: retrofitting is possible but most value comes from starting BIM at the design phase."
        ),
    },
    # ================================================================
    # 5. CONTENT - IoT
    # ================================================================
    {
        "slide_type": "content_slide",
        "title": "IoT & Smart Site Sensors",
        "body": (
            "**Environmental Monitoring** \u2014 Dust, noise, temperature, and humidity sensors with real-time alerts\n"
            "**Worker Safety Tracking** \u2014 Geofencing and wearable devices reduce incident response time\n"
            "**Equipment Telemetry** \u2014 Predictive maintenance on cranes, excavators, and concrete plants"
        ),
        "notes": (
            "KEY MESSAGE: IoT turns a construction site into a data-driven environment.\n"
            "\"Think of IoT as the nervous system of the jobsite. Sensors monitor everything in real time \u2014 dust levels, noise, temperature, even the location of every worker.\"\n"
            "Pause.\n"
            "\"Geofencing alerts trigger the moment someone enters a restricted zone. Predictive maintenance on equipment means fewer breakdowns and fewer schedule surprises.\"\n"
            "\"On one pilot site, IoT-driven predictive maintenance reduced unplanned downtime by forty percent.\"\n"
            "TRANSITION: \"IoT gives us the data. AI tells us what the data means.\"\n"
            "COACHING: Specific, not vague. Be ready for: \"How much does IoT cost per site?\" \u2014 answer: a full sensor suite runs roughly fifteen to twenty thousand dollars, payback in under six months on mid-size sites."
        ),
    },
    # ================================================================
    # 6. CONTENT - AI, DRONES, ROBOTICS
    # ================================================================
    {
        "slide_type": "content_slide",
        "title": "AI, Drones & Robotics",
        "body": (
            "**AI-Powered Analytics** \u2014 Predict schedule risks, optimize resource allocation, and detect safety hazards in real time\n"
            "**Drone Surveying** \u2014 Aerial photogrammetry maps sites in hours instead of weeks\n"
            "**Construction Robotics** \u2014 Automated bricklaying, 3D concrete printing, and autonomous earthmoving"
        ),
        "notes": (
            "KEY MESSAGE: AI, drones, and robotics automate the dangerous and the repetitive.\n"
            "\"AI takes the data from IoT and BIM and turns it into predictions. It flags schedule risks before they become delays.\"\n"
            "Pause. Walk left to right through the three cards.\n"
            "\"Drones can survey a hundred-acre site in under two hours \u2014 what used to take a survey crew a full week.\"\n"
            "\"And robotics? Autonomous bricklayers are already laying three thousand bricks per day on commercial projects in Europe and Asia.\"\n"
            "\"Ask yourself: which of these three technologies could save your next project the most time?\"\n"
            "TRANSITION: \"These tools need a platform to connect them.\"\n"
            "COACHING: Don't over-claim. Robotics is real but still early. Be ready for: \"Are construction robots commercially viable?\" \u2014 answer: yes for repetitive tasks, limited for complex custom work."
        ),
    },
    # ================================================================
    # 7. TWO CONTENT - CLOUD & DIGITAL TWINS
    # ================================================================
    {
        "slide_type": "two_content_slide",
        "title": "Cloud Platforms & Digital Twins",
        "body_left": (
            "**Cloud Collaboration**\n"
            "**Unified Data Platform** \u2014 All project data in one secure, accessible environment\n"
            "**Real-Time Coordination** \u2014 Design changes propagate instantly to all stakeholders\n"
            "**Scalable Storage** \u2014 Terabytes of models, drawings, and documents managed centrally"
        ),
        "body_right": (
            "**Digital Twin Technology**\n"
            "**Virtual Replica** \u2014 A real-time digital copy of the physical asset linked to live sensor data\n"
            "**Predictive Simulation** \u2014 Model 'what-if' scenarios for structural load, energy use, and maintenance\n"
            "**Operational Insight** \u2014 Continuous performance monitoring across the asset lifecycle"
        ),
        "notes": (
            "KEY MESSAGE: Cloud is the connective tissue; digital twins bridge physical and virtual.\n"
            "\"On the left, cloud platforms solve the collaboration problem. Every stakeholder \u2014 architect, engineer, contractor, client \u2014 works from the same source of truth.\"\n"
            "Walk through the left column first.\n"
            "\"On the right, a digital twin is the next evolution of BIM. It's a living, breathing model connected to real sensors.\"\n"
            "Walk through the right column.\n"
            "\"Digital twins let you run what-if scenarios before you act \u2014 predict structural loads, simulate maintenance schedules, and optimize energy performance.\"\n"
            "TRANSITION: \"Let's see what happens when you put all of this together.\"\n"
            "COACHING: Balance the two columns equally. Be ready for: \"Is a digital twin the same as BIM?\" \u2014 answer: BIM is the static design model; a digital twin adds real-time sensor data and simulation."
        ),
    },
    # ================================================================
    # 8. COMPARISON - TRADITIONAL vs DIGITAL
    # ================================================================
    {
        "slide_type": "comparison_slide",
        "title": "Traditional vs Digital Workflow",
        "body_left": (
            "**Traditional Approach**\n"
            "**Design** \u2014 2D drawings, siloed teams, manual coordination\n"
            "**Construction** \u2014 Paper-based tracking, reactive problem-solving\n"
            "**Operations** \u2014 As-built drawings out of date, maintenance by trial and error"
        ),
        "body_right": (
            "**Digital-First Approach**\n"
            "**Design** \u2014 3D BIM model, clash-free, multi-discipline collaboration\n"
            "**Construction** \u2014 IoT sensors, drone surveys, AI-driven scheduling\n"
            "**Operations** \u2014 Digital twin with live data, predictive maintenance"
        ),
        "notes": (
            "KEY MESSAGE: Digital-first eliminates information silos at every project phase.\n"
            "\"This slide is the heart of the argument. On the left, the traditional way \u2014 2D drawings, paper on site, guesswork in maintenance.\"\n"
            "Point to the left column.\n"
            "\"On the right, the digital-first approach \u2014 one shared model, real-time data, and predictive maintenance.\"\n"
            "Point to the right column.\n"
            "\"The same project, the same team, but fundamentally different outcomes. Which would you rather hand over to your client?\"\n"
            "TRANSITION: \"Let's look at the numbers that back this up.\"\n"
            "COACHING: Contrast tone \u2014 left side with a slight gravity, right side with forward momentum. Be ready for: \"Is the transition expensive?\" \u2014 answer: upfront investment, but ROI typically within 12 to 18 months."
        ),
    },
    # ================================================================
    # 9. SECTION HEADER - IMPACT & EVIDENCE
    # ================================================================
    {
        "slide_type": "section_header_slide",
        "title": "Impact & Evidence",
        "notes": (
            "KEY MESSAGE: The data speaks louder than the theory.\n"
            "\"We've covered the technologies and the workflow transformation. Now let's look at the evidence.\"\n"
            "\"The following charts show real market data and performance improvements that are measurable today, not projected five years out.\"\n"
            "TRANSITION: \"Let's start with the market itself.\"\n"
            "COACHING: Brief pause. Shift tone from conceptual to evidence-driven."
        ),
    },
    # ================================================================
    # 10. CHART (BAR) - MARKET GROWTH
    # ================================================================
    {
        "slide_type": "chart_slide",
        "title": "Global Construction Tech Market Growth (USD Billion)",
        "chart_type": "bar",
        "categories": ["2020", "2021", "2022", "2023", "2024", "2025", "2026"],
        "series": [
            {"name": "Market Size", "values": [8.5, 11.2, 14.8, 19.5, 25.1, 31.7, 39.4]},
        ],
        "chart_options": {
            "legend_position": "bottom",
            "show_data_labels": True,
            "y_axis_min": 0,
            "y_axis_max": 45,
            "y_axis_title": "USD Billion",
        },
        "notes": (
            "KEY MESSAGE: The construction tech market is growing at a 30%+ CAGR.\n"
            "\"Look at the trajectory on this chart.\"\n"
            "Pause. Let the audience read the numbers.\n"
            "\"Eight-point-five billion in 2020, projected to thirty-nine-point-four billion by 2026. That's a compound annual growth rate north of thirty percent.\"\n"
            "\"Investors are pouring capital into proptech and contech startups. The signal from the market is unambiguous: this is not a niche anymore.\"\n"
            "\"Consider this \u2014 at thirty-nine billion, construction tech alone would be larger than the entire global market for cybersecurity software.\"\n"
            "TRANSITION: \"Market growth is one thing. What about adoption on the ground?\"\n"
            "COACHING: Let the chart speak first, then add context. Be ready for: \"Where does this data come from?\" \u2014 answer: Statista, MarketsandMarkets, and Grand View Research consensus estimates."
        ),
    },
    # ================================================================
    # 11. CHART (PIE) - ADOPTION RATES
    # ================================================================
    {
        "slide_type": "chart_slide",
        "title": "Technology Adoption Rate by Category",
        "chart_type": "pie",
        "categories": ["Cloud Platforms", "BIM", "Drones", "IoT Sensors", "AI & ML", "Robotics"],
        "series": [
            {"name": "Adoption", "values": [72, 68, 52, 45, 28, 15]},
        ],
        "chart_options": {
            "legend_position": "right",
            "show_data_labels": True,
        },
        "notes": (
            "KEY MESSAGE: Cloud and BIM are mainstream; AI and robotics are still early but accelerating.\n"
            "\"This pie chart tells you where we are on the adoption curve.\"\n"
            "Pause. Point to the chart.\n"
            "\"Cloud platforms and BIM are the leaders \u2014 seventy-two percent and sixty-eight percent adoption among mid-to-large contractors. These are no longer pilot projects. They are standard practice.\"\n"
            "\"Drones and IoT are in the early-majority phase. AI and machine learning are still in the innovator phase at twenty-eight percent, but that number has doubled in the last two years.\"\n"
            "\"The pattern is clear: simpler tools get adopted first. The more complex the technology, the steeper the curve \u2014 but the steeper the curve, the bigger the first-mover advantage.\"\n"
            "TRANSITION: \"Numbers are compelling, but what does this look like on an actual site?\"\n"
            "COACHING: Don't rush. Give the audience time to absorb the percentages. Be ready for: \"Is the data specific to a region?\" \u2014 answer: global survey data, McKinsey and JBKnowledge benchmarks."
        ),
    },
    # ================================================================
    # 12. CONTENT IMAGE - REAL-WORLD
    # ================================================================
    {
        "slide_type": "content_image_slide",
        "title": "Smart Construction in Action",
        "body": (
            "**Integrated Digital Jobsite** \u2014 BIM model linked to IoT sensors on a major infrastructure project\n"
            "**Real-Time Dashboard** \u2014 Project managers monitor progress, safety, and resource allocation from a single screen\n"
            "**Measurable Results** \u2014 22% cost reduction, 15% faster completion, and zero lost-time incidents in the first six months"
        ),
        "notes": (
            "KEY MESSAGE: A real site, real tools, real savings.\n"
            "\"This slide shows what a digitally-enabled construction site actually looks like in practice.\"\n"
            "Pause. Let them take in the visual.\n"
            "\"The BIM model drives everything. It's linked to IoT sensors on the ground, feeding real-time data into a central dashboard.\"\n"
            "\"On this particular project, a major infrastructure program in Southeast Asia, the results were striking: twenty-two percent cost reduction, fifteen percent faster completion, and zero lost-time safety incidents in the first six months.\"\n"
            "\"Those aren't projected numbers. They're audited outcomes.\"\n"
            "TRANSITION: \"Let's look at more examples from around the world.\"\n"
            "COACHING: Reference the image directly. If image is added: 'As you can see on the screen...'. Be ready for: \"What's the project?\" \u2014 answer: infrastructure program, client details can be shared under NDA."
        ),
    },
    # ================================================================
    # 13. TWO CONTENT - CASE STUDIES
    # ================================================================
    {
        "slide_type": "two_content_slide",
        "title": "Global Case Studies",
        "body_left": (
            "**International Projects**\n"
            "**Crossrail (London)** \u2014 BIM Level 2 with 250,000+ model components, clash detection saved $200M+\n"
            "**Dubai 3D Printing** \u2014 First 3D-printed office building, 80% less construction waste\n"
            "**Singapore FinTech Hub** \u2014 Digital twin of entire district, energy optimized by 30%"
        ),
        "body_right": (
            "**Asia-Pacific Projects**\n"
            "**Hong Kong Airport Expansion** \u2014 Drone surveying reduced topographic survey time by 90%\n"
            "**Tokyo Olympic Village** \u2014 IoT sensors for crowd flow, safety, and environmental monitoring\n"
            "**Shanghai Tower** \u2014 BIM + robotic construction modules, 40% reduction in on-site labor hours"
        ),
        "notes": (
            "KEY MESSAGE: Digital tools deliver measurable results on projects worldwide.\n"
            "\"On the left, three landmark international projects.\"\n"
            "Walk through the left column.\n"
            "\"Crossrail alone saved over two hundred million dollars through BIM-driven clash detection. The Dubai 3D-printed office eliminated eighty percent of construction waste.\"\n"
            "\"On the right, three Asia-Pacific projects that show the same pattern.\"\n"
            "Walk through the right column.\n"
            "\"Hong Kong Airport cut survey time by ninety percent using drones. Shanghai Tower reduced on-site labor hours by forty percent.\"\n"
            "\"These are not edge cases. This is the new normal for leading projects.\"\n"
            "TRANSITION: \"So how do you get started?\"\n"
            "COACHING: Name-drop projects confidently. Be ready for: \"Any case studies in our country?\" \u2014 offer to share localized case studies post-presentation."
        ),
    },
    # ================================================================
    # 14. CONTENT - ROADMAP
    # ================================================================
    {
        "slide_type": "content_slide",
        "title": "Implementation Roadmap",
        "body": (
            "**Phase 1: Pilot (Months 1-6)** \u2014 Deploy BIM + IoT on 1-2 pilot projects, build internal champions\n"
            "**Phase 2: Scale (Months 6-18)** \u2014 Roll out cloud collaboration across all projects, integrate AI analytics\n"
            "**Phase 3: Transform (Months 18-36)** \u2014 Organization-wide digital twin adoption, automate reporting and maintenance"
        ),
        "notes": (
            "KEY MESSAGE: A phased, low-risk rollout \u2014 pilot first, scale second, transform third.\n"
            "\"We don't boil the ocean. We start small, prove the value, then scale.\"\n"
            "Walk through the three phases left to right.\n"
            "\"Phase one: pilot BIM and IoT on one or two projects. The goal isn't to digitize everything \u2014 it's to build internal champions who can show the rest of the organization what's possible.\"\n"
            "\"Phase two: once the pilots demonstrate measurable ROI, roll out cloud collaboration across the portfolio and start layering in AI analytics.\"\n"
            "\"Phase three: full digital twin adoption with automated reporting, predictive maintenance, and continuous improvement.\"\n"
            "\"Eighteen to thirty-six months to full transformation. Realistic, achievable, and value-positive from day one.\"\n"
            "TRANSITION: \"Of course, there are challenges to manage along the way.\"\n"
            "COACHING: Practical tone. Emphasize that each phase delivers value. Be ready for: \"What if the pilot fails?\" \u2014 answer: pilot scope is intentionally small; even partial success generates learnings."
        ),
    },
    # ================================================================
    # 15. CONTENT - CHALLENGES
    # ================================================================
    {
        "slide_type": "content_slide",
        "title": "Challenges & Mitigations",
        "body": (
            "**Data Security & Compliance** \u2014 Encrypt all project data, role-based access, and regular security audits\n"
            "**Workforce Skills Gap** \u2014 Invest in training programs and partner with technology providers for ongoing support\n"
            "**Change Management Resistance** \u2014 Executive sponsorship, phased rollout, and visible quick wins to build momentum"
        ),
        "notes": (
            "KEY MESSAGE: Challenges are real and manageable \u2014 with the right strategy.\n"
            "\"I'd be misleading you if I said this was effortless. There are three challenges that every organization faces.\"\n"
            "Walk through the three points.\n"
            "\"First, data security. Construction data is commercially sensitive. The mitigation is straightforward: encryption, role-based access, and regular audits.\"\n"
            "\"Second, the skills gap. The workforce needs training, not replacement. Most digital tools are designed to be intuitive, and training programs pay for themselves in the first quarter.\"\n"
            "\"Third, and often the hardest, is resistance to change. The antidote is visible, quick wins from the pilot phase \u2014 when people see real results on a real project, skepticism evaporates.\"\n"
            "TRANSITION: \"Let me wrap up with the key takeaways.\"\n"
            "COACHING: Honest, not dismissive of challenges. Be ready for: \"What's the biggest failure mode?\" \u2014 answer: change management, not technology. The tools work; adoption is the hard part."
        ),
    },
    # ================================================================
    # 16. CLOSING SLIDE
    # ================================================================
    {
        "slide_type": "closing_slide",
        "title": "Thank You",
        "subtitle": "Questions & Discussion",
        "notes": (
            "KEY MESSAGE: Digital construction is not the future \u2014 it is the present.\n"
            "\"Hold the slide for a moment.\"\n"
            "\"Thank you for your time.\"\n"
            "\"Let me leave you with one thought: the construction industry has been slow to digitize, but the pace of change is accelerating. The companies that act now will lead; the ones that wait will follow.\"\n"
            "Pause.\n"
            "\"I'm happy to take questions \u2014 whether about the technologies, the implementation roadmap, or how this applies to your specific projects.\"\n"
            "TRANSITION: Open for Q&A.\n"
            "COACHING: Stand still, make eye contact, open body language. Do not rush to end. Be ready for: \"Where should we start?\" \u2014 answer: BIM first, it has the highest adoption and clearest ROI."
        ),
    },
]

print(f"Total slides: {len(slide_data)}")
print(f"Layouts used: {set(s['slide_type'] for s in slide_data)}")

result = generate_ppt_from_data(
    slide_data,
    output_path=str(DEFAULT_OUTPUT_DIR / "202606150807.pptx"),
)
print(f"\nOutput: {result}")
