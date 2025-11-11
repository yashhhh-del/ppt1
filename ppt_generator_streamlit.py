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
    page_title="AI PowerPoint Generator",
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
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">📊 AI PowerPoint Generator</div>', unsafe_allow_html=True)
st.markdown("---")

# Sidebar for API Keys
with st.sidebar:
    st.header("⚙️ API Configuration")
    
    # Claude API Key
    claude_api_key = st.text_input("OpenRouter API Key *", type="password", help="Required: For generating presentation content")
    
    # Model selection
    model_choice = st.selectbox(
        "AI Model",
        ["Free Model (Google Gemini)", "Claude 3.5 Sonnet (Paid)"],
        help="Free model uses your limited credits wisely"
    )
    
    st.info("💡 Using OpenRouter API")
    
    # Stability API Key (Optional)
    stability_api_key = st.text_input("Stability AI API Key (Optional)", type="password", help="Optional: For AI-generated images")
    
    st.markdown("---")
    st.markdown("### 📖 How to Use")
    st.markdown("""
    1. Enter OpenRouter API key (required)
    2. Enter ANY topic you want
    3. Add Stability key for AI images (optional)
    4. Select Free or Paid model
    5. Click 'Generate' and get your PPT!
    """)
    st.markdown("---")
    st.markdown("### 🔗 Get API Keys")
    st.markdown("[OpenRouter API](https://openrouter.ai/keys)")
    st.markdown("[Stability AI](https://platform.stability.ai)")
    st.markdown("---")
    st.warning("💰 Low credits? Reduce slides or upgrade at OpenRouter")

# Function to generate content using Claude via OpenRouter
def generate_content_with_claude(api_key, topic, category, slide_count, tone, audience, key_points, model_choice):
    """Generate presentation content using AI via OpenRouter"""
    try:
        # Select model based on user choice
        if "Free" in model_choice:
            model = "google/gemini-2.0-flash-exp:free"
        else:
            model = "anthropic/claude-3.5-sonnet"
        
        # Calculate appropriate max_tokens based on slide count
        # Roughly 150 tokens per slide to be safe
        calculated_tokens = min(slide_count * 150 + 200, 1500)
        
        prompt = f"""You are an expert corporate presentation creator. Generate a detailed PowerPoint presentation structure.

Topic: {topic}
Category: {category}
Slide Count: {slide_count}
Tone: {tone}
Audience: {audience}
Additional Points: {key_points if key_points else "None"}

Return ONLY a JSON structure with this exact format:
{{
  "slides": [
    {{
      "title": "Title here (max 8 words)",
      "bullets": [],
      "image_prompt": "description for AI image generation"
    }},
    {{
      "title": "Slide title",
      "bullets": ["Bullet 1 (max 12 words)", "Bullet 2", "Bullet 3", "Bullet 4"],
      "image_prompt": "abstract visual description"
    }}
  ]
}}

Rules:
- First slide: Title only, no bullets
- Each slide needs: title, bullets (3-5 per slide), image_prompt
- Bullets must be actionable and specific to the topic
- Image prompts should be abstract, professional, no text
- Last slide should be conclusion/next steps
- Make content highly relevant to "{topic}"
- Total slides: exactly {slide_count}

Return ONLY valid JSON, no markdown, no explanation."""

        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key.strip()}",
                "Content-Type": "application/json",
            },
            json={
                "model": model,
                "max_tokens": calculated_tokens,  # Dynamic token allocation
                "messages": [
                    {"role": "user", "content": prompt}
                ]
            }
        )
        
        if response.status_code == 200:
            data = response.json()
            content_text = data["choices"][0]["message"]["content"]
            
            # Clean JSON response
            content_text = content_text.strip()
            if content_text.startswith("```json"):
                content_text = content_text[7:]
            if content_text.startswith("```"):
                content_text = content_text[3:]
            if content_text.endswith("```"):
                content_text = content_text[:-3]
            content_text = content_text.strip()
            
            import json
            slides_data = json.loads(content_text)
            return slides_data["slides"]
        else:
            error_data = response.json()
            if response.status_code == 402:
                st.error("💳 Insufficient credits! Options:")
                st.info("1. Reduce number of slides\n2. Add credits at https://openrouter.ai/settings/credits")
            else:
                st.error(f"OpenRouter API Error: {response.text}")
            return None
            
    except Exception as e:
        st.error(f"Error generating content: {str(e)}")
        return None

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
            return None
    except Exception as e:
        return None

# Function to create PowerPoint
def create_powerpoint(slides_content, theme, image_mode, stability_key, category, audience):
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
        if image_mode in ["AI", "Mixed"] and stability_key and idx > 0:
            image_prompt = slide_data.get("image_prompt", f"abstract representation of {slide_data['title']}")
            status_text.text(f"Generating AI image for slide {idx + 1}...")
            
            image_data = generate_image_stability(stability_key, image_prompt)
            
            if image_data:
                image_stream = io.BytesIO(image_data)
                left = Inches(6.5)
                top = Inches(2)
                pic = slide.shapes.add_picture(image_stream, left, top, height=Inches(4))
            
            time.sleep(1)
    
    progress_bar.progress(1.0)
    status_text.text("✅ Presentation created successfully!")
    
    return prs

# Main content
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📝 Your Topic")
    
    topic = st.text_input("Enter ANY Topic *", placeholder="e.g., Space Exploration, Cooking Recipes, Football History, Math Education...")
    
    st.caption("💡 Enter any topic - AI will create relevant slides automatically!")
    
    category = st.selectbox(
        "Category *",
        ["Business", "Pitch", "Marketing", "Technical", "Academic", "Training"]
    )
    
    slide_count = st.number_input("Number of Slides *", min_value=3, max_value=15, value=6, 
                                   help="⚠️ More slides = more tokens needed. Start with 6-8 slides.")
    
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
        ["Free Images (Unsplash)", "AI Generated (Paid)", "None"],
        help="Free Images: Automatic relevant images from Unsplash. AI: Custom generated images (requires Stability AI key)"
    )
    
    english_variant = st.selectbox(
        "English Variant *",
        ["US", "UK"]
    )

# Additional key points
st.subheader("➕ Additional Key Points (Optional)")
key_points = st.text_area(
    "Enter specific points you want to cover (one per line)",
    placeholder="- Focus on sustainability\n- Include case studies\n- Emphasize ROI",
    height=100
)

# Generate button
st.markdown("---")
st.info("✨ Enter ANY topic and click generate - AI will create custom content and images!")
generate_button = st.button("🚀 Generate PowerPoint on My Topic", use_container_width=True)

# Generation logic
if generate_button:
    if not claude_api_key:
        st.error("⚠️ Please enter your OpenRouter API key in the sidebar.")
    elif not claude_api_key.startswith("sk-or-"):
        st.error("⚠️ Invalid OpenRouter API key format. It should start with 'sk-or-'")
    elif not topic:
        st.error("⚠️ Please enter a topic for your presentation.")
    elif image_mode == "AI" and not stability_api_key:
        st.warning("⚠️ Stability AI key not provided. Generating presentation without images.")
        image_mode = "None"
    
    if claude_api_key and topic:
        with st.spinner("🤖 AI is analyzing your topic and creating custom content..."):
            try:
                # Generate content with Claude
                slides_content = generate_content_with_claude(
                    claude_api_key, topic, category, slide_count, 
                    tone, audience, key_points, model_choice
                )
                
                if slides_content:
                    st.success("✅ Content generated! Now creating PowerPoint...")
                    
                    # Create PowerPoint
                    prs = create_powerpoint(
                        slides_content, theme, image_mode, 
                        stability_api_key if image_mode == "AI" else None,
                        category, audience
                    )
                    
                    # Save to BytesIO
                    pptx_io = io.BytesIO()
                    prs.save(pptx_io)
                    pptx_io.seek(0)
                    
                    st.success("🎉 PowerPoint created successfully!")
                    
                    # Show preview
                    with st.expander("📄 Preview Slide Titles"):
                        for i, slide in enumerate(slides_content):
                            st.write(f"**Slide {i+1}:** {slide['title']}")
                    
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
                st.error(f"Details: {type(e).__name__}")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 1rem;'>
    <p>💡 <strong>Works with ANY topic!</strong> Try: "History of Pizza", "Quantum Physics", "Yoga Benefits", "Video Game Design"</p>
    <p>🎨 Add Stability AI key to generate custom images for your topic!</p>
    <p><strong>Free Model Available</strong> - No credits needed to start!</p>
    <p>⚠️ <strong>Low on credits?</strong> Start with 6 slides and upgrade at <a href="https://openrouter.ai/settings/credits">OpenRouter</a></p>
</div>
""", unsafe_allow_html=True)
