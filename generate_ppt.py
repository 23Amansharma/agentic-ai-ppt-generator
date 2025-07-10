from pptx import Presentation

prs = Presentation()

def add_title_slide(title, subtitle):
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle

def add_bullet_slide(title, bullet_points):
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    content = slide.placeholders[1]
    content.text = bullet_points[0]
    for point in bullet_points[1:]:
        content.text += f'\n{point}'

# Slide 1
add_title_slide("AGENTIC AI FOR INTELLIGENT TRAFFIC MANAGEMENT",
    "Satellite-Driven Congestion Control with Autonomous Agents\nTeam: Alok Gupta, Aman Sharma\nGoogle Agentic AI Hackathon 2025")

# Slide 2
add_bullet_slide("URBAN TRAFFIC CRISIS (Problem Statement)", [
    "Reactive systems lead to 22-min avg response time (Delhi Traffic Data)",
    "Satellite blind spots cause 68% downtime during monsoons",
    "Scalability limitations in conventional solutions",
    "AGENTIC SOLUTION HOOK: \"Autonomous agents enabling real-time adaptation to dynamic traffic patterns.\""
])

# (Add remaining slides here using add_bullet_slide...)

# Save
prs.save("Agentic_AI_Traffic_Management_Hackathon_2025.pptx")
