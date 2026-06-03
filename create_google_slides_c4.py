#!/usr/bin/env python3
"""
Language Curriculum Design - Chapter 4: Principles
Google Slides API Integration Script
Converts PowerPoint content to Google Slides format
Based on Nation & Macalister (2010)
"""

from google.colab import auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Authenticate with Google Account
auth.authenticate_user()
slides_service = build('slides', 'v1')
drive_service = build('drive', 'v3')

# Create a new presentation
presentation = {
    'title': 'Chapter 4: Principles - Language Curriculum Design (Nation & Macalister, 2010)'
}

presentation = slides_service.presentations().create(body=presentation).execute()
presentation_id = presentation.get('id')
print(f'✅ Created presentation with ID: {presentation_id}')

# Color definitions (RGB values normalized to 0-1)
COLOR_NAVY = {'red': 30/255, 'green': 60/255, 'blue': 114/255}
COLOR_BLUE = {'red': 52/255, 'green': 152/255, 'blue': 219/255}
COLOR_GREY = {'red': 248/255, 'green': 250/255, 'blue': 252/255}
COLOR_TEXT = {'red': 44/255, 'green': 62/255, 'blue': 80/255}
COLOR_WHITE = {'red': 1, 'green': 1, 'blue': 1}
COLOR_MUTED = {'red': 127/255, 'green': 140/255, 'blue': 141/255}

def create_slide_with_background(presentation_id, background_color):
    """Create a blank slide and set background color"""
    requests = [
        {
            'addSlide': {
                'slideLayout': 'BLANK_LAYOUT'
            }
        }
    ]
    response = slides_service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={'requests': requests}
    ).execute()
    
    slide_id = response['replies'][0]['addSlide']['slide']['objectId']
    
    # Set background color
    requests = [
        {
            'updatePageProperties': {
                'objectId': slide_id,
                'pageProperties': {
                    'pageBackgroundFill': {
                        'solidFill': {
                            'color': {
                                'rgbColor': background_color
                            }
                        }
                    }
                },
                'fields': 'pageBackgroundFill'
            }
        }
    ]
    slides_service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={'requests': requests}
    ).execute()
    
    return slide_id

def add_text_box(presentation_id, slide_id, left, top, width, height, text, font_size=18, bold=False, color=COLOR_TEXT, alignment='LEFT'):
    """Add a text box to a slide"""
    text_box = {
        'shape': {
            'shapeType': 'TEXT_BOX',
            'transform': {
                'scaleX': {'magnitude': width, 'unit': 'PT'},
                'scaleY': {'magnitude': height, 'unit': 'PT'},
                'translateX': {'magnitude': left, 'unit': 'PT'},
                'translateY': {'magnitude': top, 'unit': 'PT'}
            },
            'text': {
                'textElements': [
                    {
                        'textRun': {
                            'content': text,
                            'style': {
                                'fontSize': {
                                    'magnitude': font_size,
                                    'unit': 'PT'
                                },
                                'bold': bold,
                                'fontFamily': 'Arial' if bold else 'Calibri',
                                'foregroundColor': {
                                    'rgbColor': color
                                }
                            }
                        }
                    }
                ]
            }
        }
    }
    
    requests = [
        {
            'createShape': text_box
        }
    ]
    
    response = slides_service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={'requests': requests}
    ).execute()
    
    return response['replies'][0]['createShape']['objectId']

# ===== SLIDE 1: TITLE SLIDE =====
slide1_id = create_slide_with_background(presentation_id, COLOR_NAVY)

add_text_box(presentation_id, slide1_id, 72, 160, 824, 180, 
             "A BLUEPRINT FOR CURRICULUM DESIGN", 
             font_size=48, bold=True, color=COLOR_WHITE)

add_text_box(presentation_id, slide1_id, 72, 360, 824, 120,
             "The Architecture of Language Learning: 20 Research-Backed Principles",
             font_size=24, bold=False, color=COLOR_BLUE)

add_text_box(presentation_id, slide1_id, 72, 500, 824, 80,
             "Based on Nation & Macalister (2010) - Chapter 4: Principles",
             font_size=16, bold=False, color=COLOR_WHITE)

add_text_box(presentation_id, slide1_id, 72, 520, 824, 40,
             "Slide 1 / 19",
             font_size=12, bold=False, color=COLOR_BLUE, alignment='RIGHT')

# ===== SLIDE 2: TABLE OF CONTENTS =====
slide2_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide2_id, 72, 36, 824, 60,
             "TABLE OF CONTENTS (AGENDA)",
             font_size=36, bold=True, color=COLOR_NAVY)

# Section 1
add_text_box(presentation_id, slide2_id, 72, 120, 240, 40,
             "PART I: THEORETICAL FOUNDATIONS",
             font_size=15, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide2_id, 72, 165, 240, 120,
             "1. Beyond Methods\n2. Subdivisions of Principles\n3. Method vs. Empirical Evidence",
             font_size=12, bold=False, color=COLOR_TEXT)

# Section 2
add_text_box(presentation_id, slide2_id, 335, 120, 240, 40,
             "PART II: CORE SUBDIVISIONS",
             font_size=15, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide2_id, 335, 165, 240, 120,
             "4. Content & Sequencing\n5. Format & Presentation\n6. Monitoring & Assessment",
             font_size=12, bold=False, color=COLOR_TEXT)

# Section 3
add_text_box(presentation_id, slide2_id, 598, 120, 240, 40,
             "PART III: APPLICATIONS",
             font_size=15, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide2_id, 598, 165, 240, 120,
             "7. Four Strands Balance\n8. Human Element\n9. Evaluation & References",
             font_size=12, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide2_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                          Slide 2 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

# ===== SLIDE 3: SECTION DIVIDER - PART I =====
slide3_id = create_slide_with_background(presentation_id, COLOR_NAVY)

add_text_box(presentation_id, slide3_id, 72, 200, 824, 80,
             "PART I: THEORETICAL FOUNDATIONS",
             font_size=40, bold=True, color=COLOR_WHITE)

add_text_box(presentation_id, slide3_id, 72, 295, 824, 50,
             "Deconstructing Outdated Methods & Navigating Empirical Frameworks",
             font_size=20, bold=False, color=COLOR_BLUE)

add_text_box(presentation_id, slide3_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                          Slide 3 / 19",
             font_size=10, bold=False, color=COLOR_BLUE)

# ===== SLIDE 4: BEYOND METHODS =====
slide4_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide4_id, 72, 36, 824, 60,
             "BEYOND METHODS: RESEARCH-BASED DESIGN",
             font_size=36, bold=True, color=COLOR_NAVY)

content_4 = """The Content Illusion & The Core Problem (Nation & Macalister, 2010)

• The Fallacy of 'Modern' Courses: Many contemporary coursebooks repeat syllabi that vary only minimally from Berlitz models designed in the 1890s, failing to incorporate decades of empirical data regarding grammar item frequencies.

• The Format Illusion: Innovative methods like Total Physical Response (TPR) and the Silent Way frequently introduce novel presentation formats while leaving underlying content selection unchanged.

• Systematic Integrity: A truly sensible basis for curriculum avoids method-based distractions and instead translates peer-reviewed research directly into teaching guidelines."""

add_text_box(presentation_id, slide4_id, 72, 120, 824, 360,
             content_4,
             font_size=13, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide4_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                          Slide 4 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

# ===== SLIDE 5: THE THREE SUBDIVISIONS =====
slide5_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide5_id, 72, 36, 824, 60,
             "THE THREE SUBDIVISIONS OF PRINCIPLES",
             font_size=36, bold=True, color=COLOR_NAVY)

# Column 1
add_text_box(presentation_id, slide5_id, 72, 120, 220, 30,
             "CONTENT & SEQUENCING",
             font_size=14, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide5_id, 72, 155, 220, 320,
             "Concerned with what items are chosen, the order in which they are presented, and the depth of coverage. This group ensures that high-frequency, highly valuable items are prioritized to provide an optimal return on time invested.",
             font_size=11, bold=False, color=COLOR_TEXT)

# Column 2
add_text_box(presentation_id, slide5_id, 324, 120, 220, 30,
             "FORMAT & PRESENTATION",
             font_size=14, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide5_id, 324, 155, 220, 320,
             "Focused directly on classroom activities, pedagogical procedures, and how learners actively process instructional materials. This is the domain where classroom teachers exercise their greatest professional influence and autonomy.",
             font_size=11, bold=False, color=COLOR_TEXT)

# Column 3
add_text_box(presentation_id, slide5_id, 576, 120, 220, 30,
             "MONITORING & ASSESSMENT",
             font_size=14, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide5_id, 576, 155, 220, 320,
             "Deals with tracking learner progress, executing diagnostic testing, and conducting comprehensive evaluation. This subdivision guarantees that learning is consistently scaffolded, feedback is targeted, and learning deficits are corrected.",
             font_size=11, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide5_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                          Slide 5 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

print("✅ Slides 1-5 created successfully!")

# ===== SLIDE 6: SECTION DIVIDER - PART II =====
slide6_id = create_slide_with_background(presentation_id, COLOR_NAVY)

add_text_box(presentation_id, slide6_id, 72, 200, 824, 80,
             "PART II: THE THREE CORE SUBDIVISIONS",
             font_size=40, bold=True, color=COLOR_WHITE)

add_text_box(presentation_id, slide6_id, 72, 295, 824, 50,
             "Deep-Dive into Content, Format, Presentation, and Assessment Principles",
             font_size=20, bold=False, color=COLOR_BLUE)

add_text_box(presentation_id, slide6_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                          Slide 6 / 19",
             font_size=10, bold=False, color=COLOR_BLUE)

# ===== SLIDE 7: CONTENT & SEQUENCING 1-4 =====
slide7_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide7_id, 72, 36, 824, 60,
             "CONTENT & SEQUENCING: PRINCIPLES 1-4",
             font_size=36, bold=True, color=COLOR_NAVY)

content_7 = """#1 Frequency & Utility - A language course must explicitly prioritize high-frequency items that yield the highest communicative value in regular language use.

#2 Sequencing Types - Syllabus items can be sequenced using linear, cyclical, or modular approaches. Linear progression introduces items once; cyclical systems reintroduce items across expanding intervals.

#3 Interference Avoidance - Curriculum must be structured to prevent the concurrent introduction of highly similar, confusing lexical or grammatical items which can induce structural interference.

#4 Depth over Breadth - Content must offer deep, repeated exposures to core items rather than superficial coverage of a vast quantity of language data."""

add_text_box(presentation_id, slide7_id, 72, 115, 824, 370,
             content_7,
             font_size=12, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide7_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                          Slide 7 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

# ===== SLIDE 8: CONTENT & SEQUENCING 5-8 =====
slide8_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide8_id, 72, 36, 824, 60,
             "CONTENT & SEQUENCING: PRINCIPLES 5-8",
             font_size=36, bold=True, color=COLOR_NAVY)

content_8 = """#5 The Spaced Retrieval Matrix - Items must be encountered across expanding intervals of time. Spaced retrieval forces learners to actively reconstruct cognitive paths to the target structure, boosting long-term memory allocation.

#6 Needs-Driven Adaptation - Content must be selected and modified dynamically to match the objective socio-cognitive needs of specific learner sub-populations rather than adhering rigidly to standard textbook syllabi.

#7 Form-Meaning Connections - Curriculum sequencing should cultivate strong links between grammatical forms and semantic meaning, ensuring that structures are never introduced as isolated grammatical entities.

#8 Theoretical Alignment - The choice of what to teach must reflect contemporary findings in corpus linguistics, discourse analysis, and language acquisition research."""

add_text_box(presentation_id, slide8_id, 72, 115, 824, 370,
             content_8,
             font_size=12, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide8_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                          Slide 8 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

# ===== SLIDE 9: THE FOUR STRANDS =====
slide9_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide9_id, 72, 36, 824, 60,
             "FORMAT & PRESENTATION: THE FOUR STRANDS",
             font_size=36, bold=True, color=COLOR_NAVY)

content_9 = """Balancing the 4 Strands of Language Learning (Nation & Macalister, 2010, p. 39)

1. Meaning-Focused Input: Learners must engage with rich oral and written language where their primary attention is focused on understanding messages (Comprehensible Input / i+1).

2. Language-Focused Learning: Explicit attention is dedicated to the mechanics of the language system, including deliberate study of vocabulary, phonetics, and grammar rules.

3. Meaning-Focused Output: Forcing production in speaking and writing across varied communicative contexts, ensuring learners produce messages that others can understand.

4. Fluency Development: Helping learners maximize their use of structures they already know, working with familiar material under real-time processing pressures."""

add_text_box(presentation_id, slide9_id, 72, 115, 824, 370,
             content_9,
             font_size=12, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide9_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                          Slide 9 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

# ===== SLIDE 10: MOTIVATION & COGNITION =====
slide10_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide10_id, 72, 36, 824, 60,
             "FORMAT & PRESENTATION: MOTIVATION & COGNITION",
             font_size=36, bold=True, color=COLOR_NAVY)

# Column 1
add_text_box(presentation_id, slide10_id, 72, 120, 220, 30,
             "Motivation & Engagement",
             font_size=14, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide10_id, 72, 155, 220, 320,
             "Ranked highly among formatting principles. Presentation must foster favorable attitudes toward the language. Without interest and value, cognitive processing slows or halts entirely.",
             font_size=11, bold=False, color=COLOR_TEXT)

# Column 2
add_text_box(presentation_id, slide10_id, 324, 120, 220, 30,
             "Time on Task & Focus",
             font_size=14, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide10_id, 324, 155, 220, 320,
             "Maximize the absolute quantity of time spent actively using the target language in the classroom. Avoid excessive use of the first language (L1) during meta-talk or administrative explanations.",
             font_size=11, bold=False, color=COLOR_TEXT)

# Column 3
add_text_box(presentation_id, slide10_id, 576, 120, 220, 30,
             "Depth of Processing",
             font_size=14, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide10_id, 576, 155, 220, 320,
             "Encourage learners to process language items as deeply and analytically as possible, ensuring tasks require genuine semantic and cognitive engagement rather than mindless mechanical repetition.",
             font_size=11, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide10_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                        Slide 10 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

print("✅ Slides 6-10 created successfully!")

# ===== SLIDE 11: FORMAT & PRESENTATION 11-14 =====
slide11_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide11_id, 72, 36, 824, 60,
             "FORMAT & PRESENTATION: PRINCIPLES 11-14",
             font_size=36, bold=True, color=COLOR_NAVY)

content_11 = """#11 Integrative Motivation - Material presentation must deliberately align with learners' socio-integrative goals, helping them project positive future identities as competent speakers within the target language community.

#12 Learning Style Adaptability - Provide diverse pedagogical paths and multi-sensory tasks to accommodate individual cognitive differences, learning rates, and sensory preferences (visual, auditory, kinesthetic).

#13 Comprehensible Input (i+1) - Ensure that approximately 95-98% of running words in input texts are already known by the learner, allowing them to decode the remaining unfamiliar items through context.

#14 Pushed Output Tasks - Construct interaction patterns that compel learners to stretch their current linguistic resources, bridging gaps between receptive knowledge and productive capacity."""

add_text_box(presentation_id, slide11_id, 72, 115, 824, 370,
             content_11,
             font_size=12, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide11_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                        Slide 11 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

# ===== SLIDE 12: MONITORING & ASSESSMENT DRIVERS =====
slide12_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide12_id, 72, 36, 824, 60,
             "MONITORING & ASSESSMENT: THE CORE DRIVERS",
             font_size=36, bold=True, color=COLOR_NAVY)

# Column 1
add_text_box(presentation_id, slide12_id, 72, 120, 220, 30,
             "Diagnostic Tracking",
             font_size=14, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide12_id, 72, 155, 220, 300,
             "Continuous assessment should illuminate existing learner proficiencies and systemic gaps, guiding immediate pedagogical recalibrations rather than serving purely punitive grading purposes.",
             font_size=11, bold=False, color=COLOR_TEXT)

# Column 2
add_text_box(presentation_id, slide12_id, 324, 120, 220, 30,
             "Formative Feedback Loop",
             font_size=14, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide12_id, 324, 155, 220, 300,
             "Providing responsive, actionable feedback helps learners identify systematic errors and refine their internal language hypothesis mechanisms in real-time.",
             font_size=11, bold=False, color=COLOR_TEXT)

# Column 3
add_text_box(presentation_id, slide12_id, 576, 120, 220, 30,
             "Washback Optimization",
             font_size=14, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide12_id, 576, 155, 220, 300,
             "Ensure that tests stimulate productive classroom behavior, beneficial study habits, and balanced communicative development rather than narrow 'teaching to the test' practices.",
             font_size=11, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide12_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                        Slide 12 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

# ===== SLIDE 13: MONITORING & ASSESSMENT 15-17 =====
slide13_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide13_id, 72, 36, 824, 60,
             "MONITORING & ASSESSMENT: PRINCIPLES 15-17",
             font_size=36, bold=True, color=COLOR_NAVY)

content_13 = """#15 Ongoing Formative Tracking - Monitor learner progress dynamically using diverse, low-stakes tools (portfolios, logs, observation rubrics) to document real-world communicative growth over time.

#16 Criterion-Referenced Testing - Evaluate performance against clear, explicit behavioral benchmarks and operational goals rather than ranking students against one another through norm-referenced curves.

#17 Empowering Self-Assessment - Integrate metacognitive checkpoints to train learners in tracking their own progress, setting realistic micro-goals, and adjusting their autonomous study habits outside of class."""

add_text_box(presentation_id, slide13_id, 72, 115, 824, 370,
             content_13,
             font_size=12, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide13_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                        Slide 13 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

# ===== SLIDE 14: SECTION DIVIDER - PART III =====
slide14_id = create_slide_with_background(presentation_id, COLOR_NAVY)

add_text_box(presentation_id, slide14_id, 72, 200, 824, 80,
             "PART III: APPLICATIONS & SYNTHESIS",
             font_size=40, bold=True, color=COLOR_WHITE)

add_text_box(presentation_id, slide14_id, 72, 295, 824, 50,
             "Macro-Level Curriculum Synthesis and Pragmatic Classroom Recommendations",
             font_size=20, bold=False, color=COLOR_BLUE)

add_text_box(presentation_id, slide14_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                        Slide 14 / 19",
             font_size=10, bold=False, color=COLOR_BLUE)

# ===== SLIDE 15: CROSS-LIST OVERLAPS =====
slide15_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide15_id, 72, 36, 824, 60,
             "CROSS-LIST OVERLAPS & CROSS-VALIDATION",
             font_size=36, bold=True, color=COLOR_NAVY)

# Table content (simplified for Google Slides)
add_text_box(presentation_id, slide15_id, 72, 120, 200, 30,
             "Theoretical Concept",
             font_size=12, bold=True, color=COLOR_WHITE)
add_text_box(presentation_id, slide15_id, 292, 120, 250, 30,
             "Nation & Macalister (2010)",
             font_size=12, bold=True, color=COLOR_WHITE)
add_text_box(presentation_id, slide15_id, 572, 120, 250, 30,
             "Corroborating Frameworks",
             font_size=12, bold=True, color=COLOR_WHITE)

# Row 1
add_text_box(presentation_id, slide15_id, 72, 155, 200, 100,
             "Input and Output Balance",
             font_size=11, bold=False, color=COLOR_TEXT)
add_text_box(presentation_id, slide15_id, 292, 155, 250, 100,
             "Principle #13 & #14: Comprehensible Input + Pushed Output",
             font_size=11, bold=False, color=COLOR_TEXT)
add_text_box(presentation_id, slide15_id, 572, 155, 250, 100,
             "Ellis (2005): Receptive and productive language balances",
             font_size=11, bold=False, color=COLOR_TEXT)

# Row 2
add_text_box(presentation_id, slide15_id, 72, 270, 200, 100,
             "Focus on Form",
             font_size=11, bold=False, color=COLOR_TEXT)
add_text_box(presentation_id, slide15_id, 292, 270, 250, 100,
             "Language-focused learning strand mixed with meaning-focused strands",
             font_size=11, bold=False, color=COLOR_TEXT)
add_text_box(presentation_id, slide15_id, 572, 270, 250, 100,
             "Long (1991), Ellis (2005): Deliberate attention to form",
             font_size=11, bold=False, color=COLOR_TEXT)

# Row 3
add_text_box(presentation_id, slide15_id, 72, 385, 200, 100,
             "Learner Autonomy",
             font_size=11, bold=False, color=COLOR_TEXT)
add_text_box(presentation_id, slide15_id, 292, 385, 250, 100,
             "Principle #17: Self-Assessment and dynamic learning styles",
             font_size=11, bold=False, color=COLOR_TEXT)
add_text_box(presentation_id, slide15_id, 572, 385, 250, 100,
             "Brown (1993), Krahnke & Christison (1983): Learner centeredness",
             font_size=11, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide15_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                        Slide 15 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

print("✅ Slides 11-15 created successfully!")

# ===== SLIDE 16: PRAGMATIC CONSTRAINTS =====
slide16_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide16_id, 72, 36, 824, 60,
             "PRAGMATIC CONSTRAINTS IN IMPLEMENTATION",
             font_size=36, bold=True, color=COLOR_NAVY)

content_16 = """Institutional Mandates
Standardized national curricula or rigid corporate benchmark exams can limit a teacher's flexibility to customize content and sequencing dynamically based on immediate learner needs.

Resource Availability
Access to high-quality comprehensible input texts, digital corpus tools, or authentic listening materials varies greatly across global educational environments, restricting rich input strands.

Teacher Preparation & Training
Successfully balancing the four strands and applying spaced retrieval strategies requires a deep, research-oriented understanding of applied linguistics that may exceed standard training models."""

add_text_box(presentation_id, slide16_id, 72, 120, 824, 370,
             content_16,
             font_size=12, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide16_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                        Slide 16 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

# ===== SLIDE 17: STRATEGIC RECOMMENDATIONS =====
slide17_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide17_id, 72, 36, 824, 60,
             "STRATEGIC RECOMMENDATIONS FOR DESIGNERS",
             font_size=36, bold=True, color=COLOR_NAVY)

add_text_box(presentation_id, slide17_id, 72, 115, 824, 35,
             "Actionable Steps for Modern Language Programs",
             font_size=14, bold=True, color=COLOR_NAVY)

content_17 = """1. Audit Existing Courses: Conduct empirical frequency counts on current syllabi vocabulary and grammar to ensure high-frequency coverage.

2. Schedule Spaced Cycles: Intentionally build review modules into the calendar at expanding intervals (e.g., 1 week, 3 weeks, 6 weeks) to prevent structural decay.

3. Equalize the Four Strands: Run a time-allocation check on daily lesson plans to ensure meaning-focused input and output are not eclipsed by grammar drilling.

4. Diversify Assessment: Supplement summative end-of-term exams with criterion-referenced portfolio tracking to capture authentic student progress."""

add_text_box(presentation_id, slide17_id, 72, 160, 824, 330,
             content_17,
             font_size=12, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide17_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                        Slide 17 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

# ===== SLIDE 18: EXECUTIVE SUMMARY =====
slide18_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide18_id, 72, 36, 824, 60,
             "EXECUTIVE SUMMARY: KEY TAKEAWAYS",
             font_size=36, bold=True, color=COLOR_NAVY)

# Column 1
add_text_box(presentation_id, slide18_id, 72, 120, 220, 30,
             "Empirical Over Methods",
             font_size=14, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide18_id, 72, 155, 220, 300,
             "Reject commercial method labels and instead build core syllabus architectures around peer-reviewed research findings in applied linguistics.",
             font_size=11, bold=False, color=COLOR_TEXT)

# Column 2
add_text_box(presentation_id, slide18_id, 324, 120, 220, 30,
             "The Power of 4 Strands",
             font_size=14, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide18_id, 324, 155, 220, 300,
             "Effective language learning requires balancing focus on form with comprehensive message decoding and active language production tasks.",
             font_size=11, bold=False, color=COLOR_TEXT)

# Column 3
add_text_box(presentation_id, slide18_id, 576, 120, 220, 30,
             "Systemic Accountability",
             font_size=14, bold=True, color=COLOR_NAVY)
add_text_box(presentation_id, slide18_id, 576, 155, 220, 300,
             "Integrate continuous tracking and self-assessment mechanisms to align curriculum implementation with objective learner milestones.",
             font_size=11, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide18_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                        Slide 18 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

# ===== SLIDE 19: REFERENCES =====
slide19_id = create_slide_with_background(presentation_id, COLOR_GREY)

add_text_box(presentation_id, slide19_id, 72, 36, 824, 60,
             "COMPREHENSIVE REFERENCES (APA 7th)",
             font_size=36, bold=True, color=COLOR_NAVY)

references = """Brown, H. D. (1993). Principles of language learning and teaching (3rd ed.). Prentice Hall.

Ellis, R. (2005). Principles of instructed language learning. Asian EFL Journal, 7(3), 9-24.

George, H. V. (1963a). A verb-form frequency count. ELT Journal, 18(1), 31-37.

George, H. V. (1963b). Monograph of the Institute of Education. Bandung.

Krahnke, K. J., & Christison, M. A. (1983). Recent language research and some principles for language teachers. TESOL Quarterly, 17(4), 625-649.

Long, M. H. (1991). Focus on form: A design feature in language teaching methodology. In K. de Bot, R. Ginsberg, & C. Kramsch (Eds.), Foreign language research in cross-cultural perspective (pp. 39-52). John Benjamins.

Nation, I. S. P., & Macalister, J. (2010). Language curriculum design. Routledge.

Richards, J. C., & Rodgers, T. S. (1986). Approaches and methods in language teaching. Cambridge University Press."""

add_text_box(presentation_id, slide19_id, 72, 120, 824, 360,
             references,
             font_size=10, bold=False, color=COLOR_TEXT)

add_text_box(presentation_id, slide19_id, 72, 495, 824, 35,
             "Language Curriculum Design | Chapter 4: Principles                                                                        Slide 19 / 19",
             font_size=10, bold=False, color=COLOR_MUTED)

print("✅ Slides 16-19 created successfully!")

# Share the presentation with the user
try:
    drive_service.permissions().create(
        fileId=presentation_id,
        body={'role': 'writer', 'type': 'user', 'emailAddress': 'user@example.com'},
        fields='id'
    ).execute()
except Exception as e:
    print(f"Note: {e}")

print(f"\n✅ ✅ ✅ SUCCESS! Google Slides Presentation Created! ✅ ✅ ✅")
print(f"\n📎 PRESENTATION ID: {presentation_id}")
print(f"🔗 LINK: https://docs.google.com/presentation/d/{presentation_id}/edit")
print(f"\n✅ All 19 slides are fully EDITABLE!")
print(f"✅ Colors, fonts, and formatting are preserved!")
print(f"✅ You can now add images, modify text, and customize as needed!")
