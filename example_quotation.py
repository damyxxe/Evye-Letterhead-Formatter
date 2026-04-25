"""
Example: Generate a Quotation PDF.
Matches the provided ADE Group quotation template.
"""

from generator import generate_pdf

data = {
    "date": "9 April 2025",
    "sidebar_extra": "Quote Valid for 14 Days\nafter this date",

    # Cover page
    "client_name": "ADE Group",
    "project_title": "Digital Screen",
    "project_subtitle": "New Brand Development",

    # Phases
    "phases": [
        {
            "title": "Market Opportunity Research",
            "description": "A.I-driven Sentiment Analytics & Research (approx 15 slides).",
            "line_items": [
                {
                    "name": "Public Perception Analysis",
                    "description": (
                        "Usage of keywords to discover general perceptions of users (up to 3 markets) about "
                        "ADE Group (or market leaders/competitors) and related insights including and not "
                        "limited to: Business & Digitalisation, market growth and expansion for Singapore "
                        "Businesses, Emerging trends, etc."
                    ),
                },
                {
                    "name": "Data Compilation",
                    "description": "Organisation and key visualisation of data derived for market segmentation purposes.",
                    "fee": "Included",
                },
                {
                    "name": "Keyword Analysis",
                    "description": (
                        "Trending keywords, interests and conversations which could be used to build the "
                        "future branding narrative for the new brand & approach."
                    ),
                },
                {
                    "name": "Key Insights",
                    "description": "Insights to assist in persona-building for new Brand.",
                },
            ],
            "subtotal": "$ 10,000",
        },
        {
            "title": "Brand Strategy & Development",
            "description": (
                "Our signature BrandPilot\u2122 Programme is a 2-day hands-on and immersive brand-building and skills "
                "development workshop that aims to accomplish 3-months of brand-building work within a focussed and "
                "deliberate 2-days. The workshop will cover:"
            ),
            "line_items": [
                {
                    "name": "Brand Identity Development",
                    "description": (
                        "Consolidation and re-articulation of the business, brand, market and industry. Identify/"
                        "Consolidate/Articulate goals and objectives as a brand in the long term."
                    ),
                    "fee": "Included",
                },
                {
                    "name": "Corporate Identity System (CIS) and Messaging",
                    "description": (
                        "Inclusive of logo design or refresh, colour choices and proportions, typography and "
                        "pairing, visual identity elements, and brand application rules."
                    ),
                    "fee": "Included",
                },
                {
                    "name": "Brand Performance",
                    "description": (
                        "Identifying and articulating Unique Selling Points (USP), Performance Messaging, tone "
                        "of voice, values, purpose, and overall brand copy and content direction."
                    ),
                    "fee": "Included",
                },
                {
                    "name": "Effective Application of Corporate Identity System",
                    "description": (
                        "To apply CIS effectively and consistently for organic recognition, foster brand goodwill "
                        "and increase product value to creatively streamline marketing expenditure (for "
                        "example: application to social media, website design, presentation decks, or others)."
                    ),
                    "fee": "Included",
                },
                {
                    "name": "Brand Guidelines",
                    "description": "Digital Brand Guidelines for future applications and longevity.",
                    "fee": "Included",
                },
            ],
            "subtotal": "$ 20,000",
        },
        {
            "title": "BrandED: Brand Education & Application\nTraining",
            "description": (
                "Ensures that the strategies developed during Phase 1 and 2 endure beyond the end of the project through "
                "the creation of key collaterals, content libraries, and training of key staff."
            ),
            "line_items": [
                {
                    "name": "Content Library",
                    "description": "Developing and training to build a robust content library online to bolster marketing and create brand trust:",
                    "fee": "Included",
                    "bullets": [
                        "Engaging company history/story.",
                        "Content to attract partnerships and media enquiries.",
                        "Developing effective aftersales support.",
                        "Presenting effective and inspiring use-cases of the B2B product",
                    ],
                },
                {
                    "name": "Effective Use of B2B Engagement Tools",
                    "fee": "Included",
                    "bullets": [
                        "Development and best practices for LinkedIn presence and B2B engagement.",
                        "Creating winning Pitch/Presentation decks for B2B sales and marketing.",
                    ],
                },
                {
                    "name": "Development of \u2018Phygital\u2019 Marketing Suite",
                    "description": "Complete assessment and prioritisation of marketing collaterals needed to bring brand to market.",
                    "fee": "Included",
                    "bullets": [
                        "Planning for brochures, flyers, decks (physical or digital) as appropriate",
                        "Design across all materials for synergistic presentation of brand and content.",
                        "Training on effective use of collaterals as marketing and branding tools",
                        "Printing and production costs are not included.",
                    ],
                },
                {
                    "name": "Online Presence, Content and UI/UX Design",
                    "description": "Development of online presence and framework, including training for the consistent use of all the following:",
                    "fee": "Included",
                    "bullets": [
                        "Recommendation for best digital platform to build brand\u2019s web presence.",
                        "Design optimal user experience on chosen platform to maximise engagement",
                        "Sitemap, UI/UX, and on-brand visual and photography direction. (Photography services may be separately chargeable).",
                        "Content Flow and Content Design application for ongoing and future content.",
                        "Copywriting suggestions to improve content that\u2019s written by client, and how to maintain consistency over time.",
                    ],
                },
                {
                    "name": "Brand Education and Training",
                    "description": "Hands-on training for Client\u2019s appointed brand executive(s) on design, brand management and brand guardianship best practices and workflows",
                    "fee": "Included",
                    "bullets": [
                        "A minimum of two 1-hour sessions with appointed personnel.",
                        "Scenario exercises and brand application framework.",
                        "Skills assessment and evaluation.",
                        "Brand & marketing consultation sessions.",
                    ],
                },
            ],
            "subtotal": "$ 20,000",
        },
        {
            "title": "Brand Activation and Performance\nMarketing Framework",
            "description": (
                "The following scope are projected to commence only after the completion of the key BrandPilot\u2122 "
                "deliverables, and will last for a period of 3 months. This period may continue beyond the scope of the EDG "
                "project."
            ),
            "line_items": [
                {
                    "name": "Search Engine Optimisation (SEO) Approach",
                    "fee": "$ 1,500",
                    "bullets": [
                        "How to conduct effective Keyword Analysis",
                        "Best practices for optimising keyword outreach based on returning data over time",
                        "Efficiently designing and developing SEO-tailored content using AI tools",
                    ],
                },
                {
                    "name": "Search Engine Marketing (SEM) Strategy",
                    "fee": "$ 1,500",
                    "bullets": [
                        "How to apply base SEO principles for keyword optimisation",
                        "How to identify and allocate budgets effectively to market optimised keywords.",
                        "Framework to position future web advertising to align with identified keywords whilst preserving brand integrity.",
                    ],
                },
                {
                    "name": "B2B Social Media Engagement Strategy",
                    "description": "Analysis and identifying of best social media platform to achieve the brand\u2019s goals.",
                    "fee": "$ 3,000",
                    "bullets": [
                        "Content Calendar: How to produce, manage and maintain social media content consistently across chosen platforms based on allocated budgets.",
                        "Evaluation of social media presence and recommendations for brand launch.",
                    ],
                },
                {
                    "name": "Brand Activation and Acquisition Strategy",
                    "fee": "Included",
                    "bullets": [
                        "Brand rollout plan and activation strategy to maximise reach and impact of new brand launch to the public.",
                        "Roadmap and best practices to apply CIS, OCMS and content design towards acquiring ideal customers, partners and relevant stakeholders.",
                        "Includes design and direction of publicity rollout and brand launch management protocols",
                    ],
                },
            ],
            "discount": {"label": "Package Discount", "amount": "$ (1,000)"},
            "subtotal_label": "Fee per month",
            "subtotal": "$ 5,000",
            "total_label": "Total (3-month period)",
            "total": "$ 15,000",
        },
    ],

    # Cost summary
    "cost_summary": {
        "title": "BrandPilot + BrandED\nPackage Cost Breakdown",
        "description": (
            "This package is proposed for submission for 50% subsidy under the Enterprise Development Grant for "
            "Strategic Branding and Marketing."
        ),
        "line_items": [
            {"name": "Market Opportunity Research", "fee": "$ 10,000"},
            {"name": "Brand Strategy & Development", "fee": "$ 20,000"},
            {"name": "BrandED: Brand Education & Application Training", "fee": "$ 20,000"},
            {"name": "Brand Activation and Performance Marketing Framework", "fee": "$ 15,000"},
        ],
        "total": "$ 65,000",
        "note": "Fees quoted are in Singapore Dollars. No GST is applicable.",
    },

    # Payment milestones
    "payment_milestones": [
        {
            "amount": "30,000",
            "description": (
                "(Thirty Thousand Singapore Dollars) being 100% of the Total Fee for "
                "Market Opportunity Research and Brand Strategy and Development, payable upon "
                "<em>confirmation of Contract</em>."
            ),
        },
        {
            "amount": "20,000",
            "description": (
                "(Twenty Thousand Singapore Dollars) being 100% of the Total Fee for "
                "BrandED: Brand Education & Application Training, payable upon <em>delivery of Brand Guidelines "
                "and Corporate Identity System</em>."
            ),
        },
        {
            "amount": "5,000",
            "description": (
                "(Five Thousand Singapore Dollars) being 33.3% of the Total Fee for <em>Brand "
                "Activation and Performance Marketing Framework</em>, for the <u>first month</u> of the three months "
                "completion period, payable monthly in advance."
            ),
        },
        {
            "amount": "5,000",
            "description": (
                "(Five Thousand Singapore Dollars) being 33.3% of the Total Fee for <em>Brand "
                "Activation and Performance Marketing Framework</em>, for the <u>second month</u> of the three months "
                "completion period, payable monthly in advance."
            ),
        },
        {
            "amount": "5,000",
            "description": (
                "(Five Thousand Singapore Dollars) being 33.3% of the Total Fee for <em>Brand "
                "Activation and Performance Marketing Framework</em>, for the <u>third month</u> of the three months "
                "completion period, payable monthly in advance."
            ),
        },
    ],

    # Signatory
    "signatory_name": "Michael Ryan Chan",
    "signatory_title": "Managing Partner",

    "show_acceptance": True,
    "show_terms": False,  # Keeping this shorter for the example
}

if __name__ == "__main__":
    generate_pdf("quotation.html", data, "output/quotation.pdf")
