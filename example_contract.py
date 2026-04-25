"""
Example: Generate an Independent Contractor Agreement PDF.
Matches the provided template for Assistant Project Manager (Remote).
"""

from generator import generate_pdf

data = {
    "doc_title": "Independent Contractor Agreement",
    "doc_subtitle": "ASST. PROJECT MANAGER (REMOTE)",
    "date": "5 May 2025",

    # Recipient
    "recipient_salutation": "Ms",
    "recipient_name": "Crystalbelle Wong Shiang Yi",

    # Position details
    "role_title": "Assistant Project Manager",
    "location": "Remote",
    "job_type": "remote, full-time",
    "reporting_to": "Associate Manager",
    "start_date": "5 May 2025",
    "salary": "Singapore Dollars $1,800",
    "benefits": "Per company policy for remote service providers.",

    "position_description": (
        "Primarily, your role will be to support senior staff, contribute to content "
        "ideation and creation, maintain digital marketing channels of our clients, "
        "assist with the management of social media campaigns, and coordinate "
        "marketing projects alongside our partners and vendors."
    ),

    # Contract clauses (after Position which is hard-coded as clause 1)
    "clauses": [
        {
            "number": 2,
            "title": "SCOPE OF WORK & PRIMARY RESPONSIBILITIES",
            "content": (
                "In addition to any duties allocated by the Agency from time to time, your "
                "position will involve the responsibilities set out in Annex A attached."
            ),
            "subclauses": [],
        },
        {
            "number": 3,
            "title": "LOCATION",
            "subclauses": [
                {
                    "number": "3.1.",
                    "content": (
                        'Our offices are at 298 Tiong Bahru Road, Level 5, Suite 05007, Central Plaza, '
                        'Singapore 168730 (the "<span class="bold">Office</span>").'
                    ),
                },
                {
                    "number": "3.2.",
                    "content": "Your duties may be performed remotely in accordance with Clause 4 below.",
                },
            ],
        },
        {
            "number": 4,
            "title": "REMOTE WORKING MODEL",
            "subclauses": [
                {
                    "number": "4.1.",
                    "content": (
                        "The Agency supports a remote working model. You will be allowed to work "
                        "remotely from your home or anywhere in the world as long as you remain "
                        "contactable during our Working Hours provided for in Clause 5."
                    ),
                },
                {
                    "number": "4.2.",
                    "content": (
                        "The specific arrangement may be adjusted from time to time based on the "
                        "Agency's operational requirements and with the agreement of both Agency "
                        "and yourself."
                    ),
                },
                {
                    "number": "4.3.",
                    "content": (
                        "The Agency may modify this remote working model based on business "
                        "needs, regulatory requirements, or other relevant circumstances with 7 days "
                        "written notice."
                    ),
                },
            ],
        },
        {
            "number": 5,
            "title": "REMOTE HOURS OF WORK",
            "subclauses": [
                {
                    "number": "5.1.",
                    "content": (
                        "As a full-time Independent Contractor, you will be required to devote "
                        "substantially the whole of your time and attention during the Agency's "
                        "ordinary business hours to the performance of your duties under your "
                        "employment contract."
                    ),
                    "subclauses": [
                        {
                            "number": "5.1.1.",
                            "content": (
                                "The agency's Working Hours are between 10am to 6pm, Singapore "
                                "Time. You are entitled a 1-hour lunch break at any time between 12 "
                                "to 2pm at your own discretion."
                            ),
                        },
                        {
                            "number": "5.1.2.",
                            "content": "Your standard working hours will be 40 hours per week.",
                        },
                    ],
                },
                {
                    "number": "5.2.",
                    "content": (
                        "Recognising the need for flexibility, the Company operates a flexible working "
                        "hours policy. The timing of these hours will be agreed upon between you and "
                        "their immediate supervisor, subject to the operational requirements of the "
                        "Company."
                    ),
                },
                {
                    "number": "5.3.",
                    "content": (
                        "With our flexible working arrangements, such hours can be allocated and "
                        "reallocated according to needs from time to time."
                    ),
                },
                {
                    "number": "5.4.",
                    "content": (
                        "You may be required to work beyond the normal working hours to fulfil their "
                        "duties."
                    ),
                },
            ],
        },
        {
            "number": 6,
            "title": "SALARY AND BENEFITS",
            "subclauses": [
                {
                    "number": "6.1.",
                    "content": (
                        "Your monthly base salary will be as stipulated in Clause 1.2, subject to any "
                        "terms provided for during your Probationary Period."
                    ),
                },
                {
                    "number": "6.2.",
                    "content": (
                        "This salary will be paid by monthly deposit into your nominated bank or "
                        "payment account."
                    ),
                },
                {
                    "number": "6.3.",
                    "content": (
                        "You will be responsible for contributing to your local provident fund and/or "
                        "personal income taxes, where applicable."
                    ),
                },
                {
                    "number": "6.4.",
                    "content": (
                        "You will be entitled to certain Employee Benefits which are detailed in our "
                        "Central Portal or attached with this letter."
                    ),
                },
            ],
        },
        {
            "number": 7,
            "title": "PROBATIONARY PERIOD",
            "subclauses": [
                {
                    "number": "7.1.",
                    "content": (
                        "This agreement is subject to the satisfactory completion of a Probationary "
                        "Period or Internship of three (3) month(s), and are subject to the terms in "
                        "Clause 11 below."
                    ),
                },
                {
                    "number": "7.2.",
                    "content": (
                        "The Probationary Period or Internship is designed to grant the Agency time "
                        "to assess whether you are able to fulfil your contractual duties with the "
                        "Agency. It is also for you to gain sufficient exposure and experience under "
                        "guidance to better understand the expectations and demands of the role. "
                        "Unless otherwise stated or agreed upon, the following terms apply:"
                    ),
                    "subclauses": [
                        {"number": "7.2.1.", "content": "Your salary will be SGD$500 during the probationary period."},
                        {
                            "number": "7.2.2.",
                            "content": (
                                "For the first month, both you and the Agency will have bi-monthly "
                                "training and review sessions for smoother transition."
                            ),
                        },
                        {
                            "number": "7.2.3.",
                            "content": (
                                "Training and review sessions can occur in the form of observing or "
                                "participating in activities led by your supervisor."
                            ),
                        },
                    ],
                },
                {
                    "number": "7.3.",
                    "content": (
                        "The general terms of this contract can be mutually re-assessed upon "
                        "completion of the Probationary Period or Internship."
                    ),
                },
            ],
        },
        {
            "number": 8,
            "title": "LEAVE",
            "subclauses": [
                {
                    "number": "8.1.",
                    "content": (
                        "You will be entitled to paid annual leave of fourteen (14) working days each "
                        "year, subject to the Agency's policies and applicable laws, and only after "
                        "completion of the Probationary Period. Leave is generally not permissible "
                        "during the Probationary Period unless explicitly granted by written "
                        "agreement."
                    ),
                },
                {
                    "number": "8.2.",
                    "content": (
                        "You may also be entitled to sick leave, parental leave and childcare leave "
                        "subject to the Agency's policies and applicable laws and regulations."
                    ),
                },
            ],
        },
        {
            "number": 9,
            "title": "COMPANY POLICIES",
            "subclauses": [
                {
                    "number": "9.1.",
                    "content": (
                        "You agree that the Agency's policies, as amended or replaced from time to "
                        "time, shall be binding upon you, but shall not form part of the employment "
                        "contract."
                    ),
                },
                {
                    "number": "9.2.",
                    "content": (
                        "You may access these policies at any time by request, or by accessing them "
                        "on our Central Portal."
                    ),
                },
                {
                    "number": "9.3.",
                    "content": "Such policies may be modified by the Agency at any time, with notice.",
                },
            ],
        },
        {
            "number": 10,
            "title": "CONFIDENTIALITY AND INTELLECTUAL PROPERTY",
            "subclauses": [
                {
                    "number": "10.1.",
                    "content": (
                        "You agree that you will not divulge any of the confidential information or "
                        "trade secrets of the Agency to any person, whether during or after the "
                        "termination of your employment."
                    ),
                },
                {
                    "number": "10.2.",
                    "content": (
                        "You agree that you will not use, attempt to use, or assist another person in "
                        "using any confidential information you may acquire in the course of your "
                        "employment in a manner which may cause loss to the Agency."
                    ),
                },
                {
                    "number": "10.3.",
                    "content": (
                        "All Intellectual Property (IP) created while under the employ of the Agency "
                        "and in service of the Agency, shall ostensibly first belong to the Agency. This "
                        "is subject to the internal IP policies of the Agency, which can be accessed at "
                        "any time via the Central Portal."
                    ),
                },
                {
                    "number": "10.4.",
                    "content": (
                        "All personal information shall be kept confidential and protected as required "
                        "under the Singapore Personal Data Protection Act, and shall solely be used "
                        "for the purposes of employment and personnel management."
                    ),
                },
            ],
        },
        {
            "number": 11,
            "title": "TERM & TERMINATION",
            "subclauses": [
                {"number": "11.1.", "content": "This Agreement shall continue and remain in effect until terminated."},
                {
                    "number": "11.2.",
                    "content": (
                        "During the Probationary Period, either party may terminate this contract by "
                        "providing fourteen (14) day(s) of written notice (or payment in lieu of notice) "
                        "to the other party."
                    ),
                },
                {
                    "number": "11.3.",
                    "content": (
                        "On expiry of the Probationary Period, either party may terminate your "
                        "employment contract by providing thirty (30) day(s) of written notice (or "
                        "payment in lieu of notice) to the other party."
                    ),
                },
                {
                    "number": "11.4.",
                    "content": (
                        "Notwithstanding sub-clause 11.1 and 11.2 above, the Agency may terminate "
                        "your employment contract by notice effective immediately without payment "
                        "(except salary accrued to the date of termination) where you have committed "
                        "an act of wilful or serious misconduct, are significantly neglectful of your "
                        "duties, or you are in serious breach of your employment contract."
                    ),
                },
                {
                    "number": "11.5.",
                    "content": "The Agency may negotiate other terms for termination via mutual agreement.",
                },
            ],
        },
        {
            "number": 12,
            "title": "CONDITIONS OF CONTRACT",
            "subclauses": [
                {
                    "number": "12.1.",
                    "content": (
                        "You, the Contractor, is engaged as an independent contractor. Nothing "
                        "herein shall be construed to create an employer-employee relationship. The "
                        "Contractor is not entitled to any benefits provided to employees of the "
                        "Company."
                    ),
                },
                {
                    "number": "12.2.",
                    "content": "The Contractor warrants that he or she:",
                    "subclauses": [
                        {"number": "12.2.1.", "content": "Has the legal capacity and authority to enter into this agreement;"},
                        {
                            "number": "12.2.2.",
                            "content": "Will perform the Services in accordance and compliance with all applicable laws;",
                        },
                        {
                            "number": "12.2.3.",
                            "content": (
                                "Is not and shall not be engaged in corrupt practices in dealing with "
                                "third parties or for acquiring business for the Agency;"
                            ),
                        },
                        {
                            "number": "12.2.4.",
                            "content": (
                                "Does not hold any prior criminal record and is not in the midst of "
                                "any criminal investigation or civil lawsuit; and"
                            ),
                        },
                        {
                            "number": "12.2.5.",
                            "content": (
                                "Has not carried nor is carrying on any other business that may "
                                "damage the repute, standing, solvency or otherwise the integrity of "
                                "the Agency."
                            ),
                        },
                    ],
                },
                {
                    "number": "12.3.",
                    "content": (
                        "The Contractor shall hereby hold harmless the Agency and all its agents, "
                        "contractors and other connected parties against:"
                    ),
                    "subclauses": [
                        {
                            "number": "12.3.1.",
                            "content": (
                                "Any Damages incurred by, or results in a proceeding against the "
                                "Agency, arising from or connected with any breach under this "
                                "Agreement; or"
                            ),
                        },
                        {
                            "number": "12.3.2.",
                            "content": (
                                "Any liabilities of the Contractor not fully disclosed to the Agency in "
                                "writing prior to this Agreement which results in Damages to the "
                                "Agency; or"
                            ),
                        },
                        {
                            "number": "12.3.3.",
                            "content": "Any breach of warranty specified in Clause 12.2 above.",
                        },
                    ],
                },
            ],
        },
    ],

    # Signatory
    "signatory_name": "Michael Ryan Chan",
    "signatory_title": "Managing Partner",
    "signatory_date": "Monday, 5 May 2025",

    # Annex A
    "annexes": [
        {
            "title": "Annex A",
            "subtitle": "SCOPE OF WORK & PRIMARY RESPONSIBILITIES",
            "intro": (
                "This Schedule outlines the key responsibilities and expectations for the role of "
                "<span class='bold'>Assistant Project Manager</span>, to be read in conjunction with the terms of the "
                "Employment Agreement."
            ),
            "sections": [
                {
                    "number": 1,
                    "title": "PURPOSE OF ROLE",
                    "content": (
                        "You shall support the Company in the coordination, execution, and delivery "
                        "of client and internal projects across branding, design, and marketing "
                        "functions."
                    ),
                },
                {
                    "number": 2,
                    "title": "CORE RESPONSIBILITIES",
                    "subclauses": [
                        {
                            "number": "2.1.",
                            "content": (
                                "Manage client communications via WhatsApp and email during official office "
                                "hours, ensuring prompt and professional responses."
                            ),
                        },
                        {
                            "number": "2.2.",
                            "content": (
                                "Use Notion to track project timelines, deliverables, and internal workflows "
                                "across multiple client accounts."
                            ),
                        },
                        {
                            "number": "2.3.",
                            "content": (
                                "Support the development, coordination, and basic performance monitoring "
                                "of digital campaigns across platforms (e.g., Facebook, Instagram, TikTok, "
                                "LinkedIn)."
                            ),
                        },
                        {
                            "number": "2.4.",
                            "content": (
                                "Liaise with creative and marketing team members to translate client briefs "
                                "into actionable tasks."
                            ),
                        },
                        {
                            "number": "2.5.",
                            "content": (
                                "Prepare regular client reports and contribute strategic suggestions based on "
                                "performance data and feedback."
                            ),
                        },
                        {
                            "number": "2.6.",
                            "content": (
                                "Analyse client feedback, help refine strategies, and ensure client satisfaction "
                                "by suggesting optimisations and improvements."
                            ),
                        },
                        {
                            "number": "2.7.",
                            "content": (
                                "Proactively identify and resolve project delivery issues and support timely "
                                "execution of all deliverables."
                            ),
                        },
                    ],
                },
                {
                    "number": 3,
                    "title": "COLLABORATION AND TOOLS",
                    "subclauses": [
                        {
                            "number": "3.1.",
                            "content": (
                                "The Employee will work closely with the Associate Manager, Creative "
                                "Associates, and other team members."
                            ),
                        },
                        {
                            "number": "3.2.",
                            "content": (
                                "Communication and task management will be conducted via Notion, Zoom, "
                                "Telegram, and other tools designated by the Company."
                            ),
                        },
                    ],
                },
                {
                    "number": 4,
                    "title": "WORK ARRANGEMENT",
                    "subclauses": [
                        {
                            "number": "4.1.",
                            "content": (
                                "You shall work remotely and remain communicatively available during "
                                "standard office hours (10:00 AM to 6:00 PM Singapore Time), Monday to "
                                "Friday."
                            ),
                        },
                        {
                            "number": "4.2.",
                            "content": (
                                "Flexibility is provided in work hours subject to the timely completion of "
                                "assigned tasks."
                            ),
                        },
                    ],
                },
                {
                    "number": 5,
                    "title": "ADDITIONAL DUTIES",
                    "subclauses": [
                        {
                            "number": "5.1.",
                            "content": (
                                "The Employee may be required to undertake ad hoc assignments, "
                                "administrative tasks, or project support functions beyond the core "
                                "responsibilities listed above."
                            ),
                        },
                        {
                            "number": "5.2.",
                            "content": (
                                'These tasks will be within <span class="bold">reasonable scope</span> and aligned with your skill set '
                                "and availability, and shall not substantially alter the nature of your role."
                            ),
                        },
                        {
                            "number": "5.3.",
                            "content": "Ad hoc duties shall not exceed 30% of your time and scope.",
                        },
                    ],
                },
                {
                    "number": 6,
                    "title": "PERFORMANCE AND CONDUCT",
                    "subclauses": [
                        {
                            "number": "6.1.",
                            "content": (
                                "Your are expected to demonstrate a proactive, detail-oriented, and client-"
                                "focused approach."
                            ),
                        },
                        {
                            "number": "6.2.",
                            "content": (
                                "To maintain professional conduct, responsiveness, and a commitment to "
                                "collaborative teamwork are essential at all times."
                            ),
                        },
                    ],
                },
            ],
        },
    ],

    "show_acceptance": True,
}

if __name__ == "__main__":
    generate_pdf("contract.html", data, "output/contract.pdf")
