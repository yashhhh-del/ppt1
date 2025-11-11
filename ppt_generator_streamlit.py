import streamlit as st
import requests
import base64
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import time

# Page configuration
st.set_page_config(
    page_title="AI PowerPoint Generator with Stability AI",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #1f77b4;
        color: white;
        font-weight: bold;
        padding: 0.5rem;
        border-radius: 5px;
    }
    .output-box {
        background-color: #f0f2f6;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
        margin-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">📊 AI PowerPoint Generator with Stability AI</div>', unsafe_allow_html=True)
st.markdown("---")

# Sidebar for API Key
with st.sidebar:
    st.header("⚙️ Configuration")
    stability_api_key = st.text_input("Stability AI API Key", type="password", help="Enter your Stability AI API key")
    st.markdown("---")
    st.markdown("### 📖 How to Use")
    st.markdown("""
    1. Enter your Stability AI API key
    2. Fill in presentation details
    3. Click 'Generate Presentation'
    4. Download your PPT with AI images
    """)
    st.markdown("---")
    st.markdown("### 🔗 Get API Key")
    st.markdown("[Get Stability AI Key](https://platform.stability.ai/)")
    st.markdown("---")
    st.markdown("Made with ❤️ using Stability AI")

# Function to generate image using Stability AI
def generate_image_stability(api_key, prompt):
    """Generate image using Stability AI API"""
    try:
        response = requests.post(
            "https://api.stability.ai/v1/generation/stable-diffusion-xl-1024-v1-0/text-to-image",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "text_prompts": [
                    {
                        "text": f"{prompt}, professional, clean, abstract, minimal, no text, no words, no letters",
                        "weight": 1
                    }
                ],
                "cfg_scale": 7,
                "height": 512,
                "width": 512,
                "samples": 1,
                "steps": 30,
            },
        )
        
        if response.status_code == 200:
            data = response.json()
            image_data = base64.b64decode(data["artifacts"][0]["base64"])
            return image_data
        else:
            st.warning(f"Image generation failed: {response.text}")
            return None
    except Exception as e:
        st.warning(f"Error generating image: {str(e)}")
        return None

# Function to create PowerPoint
def create_powerpoint(topic, category, slide_count, tone, audience, theme, image_mode, key_points, api_key):
    """Create PowerPoint presentation"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Define theme colors
    themes = {
        "Corporate Blue": {"bg": RGBColor(240, 248, 255), "accent": RGBColor(31, 119, 180), "text": RGBColor(0, 0, 0)},
        "Gradient Modern": {"bg": RGBColor(240, 242, 246), "accent": RGBColor(138, 43, 226), "text": RGBColor(0, 0, 0)},
        "Minimal Dark": {"bg": RGBColor(30, 30, 30), "accent": RGBColor(255, 215, 0), "text": RGBColor(255, 255, 255)},
        "Pastel Soft": {"bg": RGBColor(255, 250, 240), "accent": RGBColor(255, 182, 193), "text": RGBColor(60, 60, 60)}
    }
    
    color_scheme = themes.get(theme, themes["Corporate Blue"])
    
    # Generate slide content based on category and topic
    slides_content = generate_slide_content(topic, category, slide_count, tone, audience, key_points)
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, slide_data in enumerate(slides_content):
        status_text.text(f"Creating slide {idx + 1}/{len(slides_content)}...")
        progress_bar.progress((idx + 1) / len(slides_content))
        
        # Add blank slide
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # Set background color
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = color_scheme["bg"]
        
        # Add title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = slide_data["title"]
        title_frame.paragraphs[0].font.size = Pt(36 if idx == 0 else 28)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].font.color.rgb = color_scheme["accent"]
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER if idx == 0 else PP_ALIGN.LEFT
        
        # Add bullets (if not first slide)
        if idx > 0 and slide_data.get("bullets"):
            bullet_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(5.5), Inches(4.5))
            text_frame = bullet_box.text_frame
            text_frame.word_wrap = True
            
            for bullet in slide_data["bullets"]:
                p = text_frame.add_paragraph()
                p.text = bullet
                p.level = 0
                p.font.size = Pt(18)
                p.font.color.rgb = color_scheme["text"]
                p.space_after = Pt(12)
        
        # Add subtitle on first slide
        if idx == 0:
            subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(3), Inches(9), Inches(1))
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = f"{category} Presentation | {audience}"
            subtitle_frame.paragraphs[0].font.size = Pt(20)
            subtitle_frame.paragraphs[0].font.color.rgb = color_scheme["text"]
            subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Generate and add image if needed
        if image_mode in ["AI", "Mixed"] and api_key and idx > 0:
            image_prompt = slide_data.get("image_prompt", f"abstract representation of {slide_data['title']}")
            status_text.text(f"Generating AI image for slide {idx + 1}...")
            
            image_data = generate_image_stability(api_key, image_prompt)
            
            if image_data:
                image_stream = io.BytesIO(image_data)
                left = Inches(6.5)
                top = Inches(2)
                pic = slide.shapes.add_picture(image_stream, left, top, height=Inches(4))
            
            time.sleep(1)
    
    progress_bar.progress(1.0)
    status_text.text("✅ Presentation created successfully!")
    
    return prs

# Function to generate slide content
def generate_slide_content(topic, category, slide_count, tone, audience, key_points):
    """Generate slide content structure"""
    slides = []
    
    # Slide 1: Title
    slides.append({
        "title": topic,
        "bullets": [],
        "image_prompt": None
    })
    
    # Slide 2: Overview/Agenda
    slides.append({
        "title": "Overview",
        "bullets": [
            f"Understanding the {category.lower()} landscape",
            "Key objectives and goals",
            "Strategic approach and methodology",
            "Expected outcomes and benefits"
        ],
        "image_prompt": "business strategy overview abstract"
    })
    
    # Parse key points if provided
    parsed_points = []
    if key_points:
        parsed_points = [p.strip('- ').strip() for p in key_points.split('\n') if p.strip()]
    
    # Middle slides based on category
    if category == "Business":
        middle_slides = [
            {"title": "Current Market Analysis", "bullets": ["Market size and growth trends", "Competitive landscape overview", "Customer needs and pain points", "Industry challenges and opportunities"], "image_prompt": "market analysis chart abstract"},
            {"title": "Strategic Objectives", "bullets": parsed_points[:4] if parsed_points else ["Increase market share and revenue", "Enhance operational efficiency", "Drive innovation and growth", "Strengthen brand position"], "image_prompt": "business goals target abstract"},
            {"title": "Implementation Plan", "bullets": ["Phase-wise rollout strategy", "Resource allocation and timeline", "Key milestones and deliverables", "Risk management approach"], "image_prompt": "project timeline roadmap abstract"},
        ]
    elif category == "Pitch":
        middle_slides = [
            {"title": "The Problem", "bullets": ["Current market gap identified", "Customer pain points validated", "Size of opportunity quantified", "Urgency for solution established"], "image_prompt": "problem challenge abstract"},
            {"title": "Our Solution", "bullets": parsed_points[:4] if parsed_points else ["Innovative approach to solve problem", "Unique value proposition delivered", "Competitive advantages highlighted", "Technology-driven implementation"], "image_prompt": "solution innovation abstract"},
            {"title": "Business Model", "bullets": ["Revenue streams and pricing", "Customer acquisition strategy", "Scalability and growth potential", "Unit economics and profitability"], "image_prompt": "business model revenue abstract"},
        ]
    elif category == "Marketing":
        middle_slides = [
            {"title": "Target Audience", "bullets": ["Demographics and psychographics defined", "Customer personas developed", "Pain points and motivations", "Buying behavior patterns"], "image_prompt": "target audience people abstract"},
            {"title": "Marketing Strategy", "bullets": parsed_points[:4] if parsed_points else ["Multi-channel campaign approach", "Content marketing initiatives", "Social media engagement plan", "Influencer partnerships"], "image_prompt": "marketing strategy megaphone abstract"},
            {"title": "Campaign Execution", "bullets": ["Timeline and key milestones", "Budget allocation across channels", "Performance metrics and KPIs", "Testing and optimization plan"], "image_prompt": "campaign execution calendar abstract"},
        ]
    else:
        middle_slides = [
            {"title": "Key Concepts", "bullets": parsed_points[:4] if parsed_points else ["Fundamental principles explained", "Core components identified", "System architecture overview", "Technical requirements defined"], "image_prompt": "technology concepts abstract"},
            {"title": "Methodology", "bullets": ["Step-by-step approach outlined", "Best practices and standards", "Tools and resources required", "Quality assurance measures"], "image_prompt": "process workflow abstract"},
            {"title": "Implementation", "bullets": ["Practical application examples", "Hands-on demonstrations", "Common challenges addressed", "Troubleshooting guidelines"], "image_prompt": "implementation execution abstract"},
        ]
    
    # Add middle slides up to slide_count - 2
    remaining_slots = slide_count - 2
    for i in range(min(remaining_slots - 1, len(middle_slides))):
        slides.append(middle_slides[i])
    
    # Final slide: Conclusion/CTA
    slides.append({
        "title": "Next Steps",
        "bullets": [
            "Key takeaways and action items",
            "Implementation timeline and milestones",
            "Resources and support available",
            "Contact information and follow-up"
        ],
        "image_prompt": "success achievement abstract"
    })
    
    return slides[:slide_count]

# Main content
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📝 Presentation Details")
    
    topic = st.text_input("Topic *", placeholder="e.g., Digital Marketing Strategy 2025")
    
    category = st.selectbox(
        "Category *",
        ["Business", "Pitch", "Marketing", "Technical", "Academic", "Training"]
    )
    
    slide_count = st.number_input("Number of Slides *", min_value=3, max_value=15, value=8)
    
    tone = st.selectbox(
        "Tone *",
        ["Formal", "Neutral", "Inspirational"]
    )

with col2:
    st.subheader("🎨 Style & Audience")
    
    audience = st.selectbox(
        "Audience *",
        ["Investors", "Students", "Corporate", "Clients", "Managers"]
    )
    
    theme = st.selectbox(
        "Theme Style *",
        ["Corporate Blue", "Gradient Modern", "Minimal Dark", "Pastel Soft"]
    )
    
    image_mode = st.selectbox(
        "Image Mode *",
        ["AI", "Mixed", "None"]
    )
    
    english_variant = st.selectbox(
        "English Variant *",
        ["US", "UK"]
    )

# Additional key points
st.subheader("➕ Additional Key Points (Optional)")
key_points = st.text_area(
    "Enter key points (one per line)",
    placeholder="- Increase market share by 25%\n- Focus on digital transformation\n- Reduce operational costs",
    height=100
)

# Generate button
st.markdown("---")
generate_button = st.button("🚀 Generate PowerPoint with AI Images", use_container_width=True)

# Generation logic
if generate_button:
    if not topic:
        st.error("⚠️ Please enter a topic for your presentation.")
    elif image_mode in ["AI", "Mixed"] and not stability_api_key:
        st.error("⚠️ Please enter your Stability AI API key in the sidebar for AI image generation.")
    else:
        with st.spinner("🎨 Creating your professional presentation with AI images..."):
            try:
                # Create PowerPoint
                prs = create_powerpoint(
                    topic, category, slide_count, tone, 
                    audience, theme, image_mode, key_points, 
                    stability_api_key if image_mode in ["AI", "Mixed"] else None
                )
                
                # Save to BytesIO
                pptx_io = io.BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)
                
                st.success("✅ PowerPoint created successfully with AI-generated images!")
                
                # Download button
                st.download_button(
                    label="📥 Download PowerPoint",
                    data=pptx_io,
                    file_name=f"{topic.replace(' ', '_')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
            except requests.exceptions.RequestException as e:
                st.error(f"❌ API Error: {str(e)}")
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 1rem;'>
    <p>💡 <strong>Tip:</strong> AI image generation takes time. Be patient for best results!</p>
    <p>Get your Stability AI API key from <a href='https://platform.stability.ai/' target='_blank'>platform.stability.ai</a></p>
</div>
""", unsafe_allow_html=True)
