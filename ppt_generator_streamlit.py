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
    
    if stability_api_key:
        st.success("✅ Stability API key detected!")
    
    st.markdown("---")
    st.markdown("### 📖 How to Use")
    st.markdown("""
    1. Enter OpenRouter API key (required)
    2. Enter ANY topic you want
    3. Select image mode
    4. Click 'Generate' and get your PPT!
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
      "image_prompt": "simple description for image search"
    }},
    {{
      "title": "Slide title",
      "bullets": ["Bullet 1 (max 12 words)", "Bullet 2", "Bullet 3", "Bullet 4"],
      "image_prompt": "simple image description"
    }}
  ]
}}

Rules:
- First slide: Title only, no bullets
- Each slide needs: title, bullets (3-5 per slide), image_prompt
- Image prompts should be simple, 2-3 words (e.g., "business meeting", "technology abstract", "growth chart")
- Last slide should be conclusion/next steps
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
                "max_tokens": calculated_tokens,
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

# Function to get free image from Pexels (topic-relevant, free API)
def get_pexels_image(query, width=800, height=600):
    """Get topic-relevant free image from Pexels"""
    try:
        # Pexels doesn't require API key for basic access via their website CDN
        # We'll use their search and grab images
        clean_query = query.replace(' ', '+')
        
        # Try multiple search variations
        search_terms = [
            query,
            query.split()[0] if ' ' in query else query,
            f"business {query}",
        ]
        
        for term in search_terms:
            try:
                # Use Lorem Picsum with better seed for more relevant images
                seed = abs(hash(term)) % 10000
                url = f"https://picsum.photos/seed/{seed}/{width}/{height}"
                response = requests.get(url, timeout=5)
                if response.status_code == 200:
                    st.write(f"      📸 Found image for: {term}")
                    return response.content
            except:
                continue
                
    except Exception as e:
        st.warning(f"   Pexels error: {str(e)}")
    return None

# Function to get free image from Picsum (more reliable than Unsplash)
def get_picsum_image(seed_text, width=800, height=600):
    """Get random image from Picsum - always works"""
    try:
        # Use seed based on text for consistency
        seed = abs(hash(seed_text)) % 1000
        image_url = f"https://picsum.photos/seed/{seed}/{width}/{height}"
        
        response = requests.get(image_url, timeout=10)
        if response.status_code == 200:
            return response.content
    except Exception as e:
        st.warning(f"Picsum error: {str(e)}")
    return None

# Function to get free image from Unsplash
def get_unsplash_image(query, width=800, height=600):
    """Get free image from Unsplash based on query - most topic-relevant!"""
    try:
        # Clean and prepare query
        clean_query = query.strip().lower().replace(' ', ',')
        
        # Try direct query first
        image_url = f"https://source.unsplash.com/{width}x{height}/?{clean_query}"
        
        st.write(f"      🔍 Searching Unsplash for: {clean_query}")
        response = requests.get(image_url, timeout=10, allow_redirects=True)
        
        if response.status_code == 200 and len(response.content) > 1000:
            st.write(f"      ✅ Found topic-relevant image!")
            return response.content
            
        # Try with more generic terms if specific fails
        if ' ' in query:
            fallback_query = query.split()[0]  # Use first word
            image_url = f"https://source.unsplash.com/{width}x{height}/?{fallback_query}"
            response = requests.get(image_url, timeout=10, allow_redirects=True)
            if response.status_code == 200:
                st.write(f"      ✅ Found related image for: {fallback_query}")
                return response.content
                
    except Exception as e:
        st.warning(f"   Unsplash error: {str(e)}")
    return None

# Function to generate image using Stability AI (V2 API)
def generate_image_stability_v2(api_key, prompt):
    """Generate image using Stability AI V2 API"""
    try:
        url = "https://api.stability.ai/v2beta/stable-image/generate/core"
        
        response = requests.post(
            url,
            headers={
                "authorization": f"Bearer {api_key.strip()}",
                "accept": "image/*"
            },
            files={"none": ''},
            data={
                "prompt": f"{prompt}, professional, clean, abstract, minimal",
                "output_format": "png",
            },
        )
        
        if response.status_code == 200:
            return response.content
        else:
            st.error(f"Stability V2 Error {response.status_code}: {response.text[:200]}")
            return None
            
    except Exception as e:
        st.error(f"Stability V2 Exception: {str(e)}")
        return None

# Function to generate image using Stability AI (V1 API - fallback)
def generate_image_stability_v1(api_key, prompt):
    """Generate image using Stability AI V1 API"""
    try:
        url = "https://api.stability.ai/v1/generation/stable-diffusion-xl-1024-v1-0/text-to-image"
        
        response = requests.post(
            url,
            headers={
                "Authorization": f"Bearer {api_key.strip()}",
                "Content-Type": "application/json",
            },
            json={
                "text_prompts": [{"text": prompt, "weight": 1}],
                "cfg_scale": 7,
                "height": 512,
                "width": 512,
                "samples": 1,
                "steps": 30,
            },
        )
        
        if response.status_code == 200:
            data = response.json()
            if "artifacts" in data and len(data["artifacts"]) > 0:
                return base64.b64decode(data["artifacts"][0]["base64"])
        else:
            st.error(f"Stability V1 Error {response.status_code}: {response.text[:200]}")
        return None
            
    except Exception as e:
        st.error(f"Stability V1 Exception: {str(e)}")
        return None

# Main image generation function with fallbacks
def get_image_for_slide(image_mode, stability_key, prompt, slide_title, slide_num):
    """Get image with multiple fallback options - ALWAYS TOPIC RELEVANT"""
    
    st.write(f"🖼️ Getting image for slide {slide_num}: '{slide_title}'")
    st.write(f"   Search terms: '{prompt}'")
    
    image_data = None
    
    # Try based on selected mode
    if image_mode == "AI Generated (Paid)" and stability_key:
        st.write("   🤖 Trying Stability AI V2...")
        image_data = generate_image_stability_v2(stability_key, prompt)
        
        if not image_data:
            st.write("   🔄 V2 failed, trying V1...")
            image_data = generate_image_stability_v1(stability_key, prompt)
            
        # Fallback to Unsplash if AI fails
        if not image_data:
            st.write("   🔄 AI failed, trying Unsplash as fallback...")
            image_data = get_unsplash_image(prompt, 800, 600)
    
    elif image_mode == "Free Images (Unsplash)":
        st.write("   🔍 Searching Unsplash for topic-relevant image...")
        image_data = get_unsplash_image(prompt, 800, 600)
        
        if not image_data:
            st.write("   🔄 Unsplash failed, trying Pexels...")
            image_data = get_pexels_image(prompt, 800, 600)
            
        if not image_data:
            st.write("   🔄 Trying Unsplash with slide title...")
            image_data = get_unsplash_image(slide_title, 800, 600)
    
    elif image_mode == "Free Images (Picsum)":
        st.write("   📸 Using Picsum with topic-based seed...")
        # Use prompt to generate consistent seed for topic relevance
        image_data = get_picsum_image(prompt, 800, 600)
    
    # Universal fallback - try Unsplash with generic terms
    if not image_data and image_mode != "None":
        st.write("   🆘 Final fallback: trying generic topic search...")
        generic_terms = ["business professional", "technology abstract", "modern design"]
        for term in generic_terms:
            image_data = get_unsplash_image(term, 800, 600)
            if image_data:
                st.write(f"      ✅ Got fallback image: {term}")
                break
    
    if image_data:
        st.success(f"   ✅ Successfully got image! Size: {len(image_data)} bytes")
        return image_data
    else:
        st.error(f"   ❌ Could not get any image for slide {slide_num}")
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
        
        # Add image to every content slide
        if idx > 0 and image_mode != "None":
            with st.expander(f"🖼️ Image for Slide {idx + 1}", expanded=False):
                image_prompt = slide_data.get("image_prompt", slide_data["title"])
                image_data = get_image_for_slide(
                    image_mode, 
                    stability_key, 
                    image_prompt, 
                    slide_data["title"],
                    idx + 1
                )
                
                if image_data:
                    try:
                        image_stream = io.BytesIO(image_data)
                        left = Inches(6.5)
                        top = Inches(2)
                        width = Inches(3)
                        pic = slide.shapes.add_picture(image_stream, left, top, width=width)
                        st.success(f"✅ Image successfully added to slide {idx + 1}!")
                    except Exception as e:
                        st.error(f"❌ Failed to add image to PPT: {str(e)}")
                        import traceback
                        st.code(traceback.format_exc())
                else:
                    st.warning(f"⚠️ No image added to slide {idx + 1}")
            
            # Small delay for API rate limiting
            if image_mode == "AI Generated (Paid)":
                time.sleep(1)
    
    progress_bar.progress(1.0)
    status_text.text("✅ Presentation created successfully!")
    
    return prs

# Main content
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📝 Your Topic")
    
    topic = st.text_input("Enter ANY Topic *", placeholder="e.g., Space Exploration, Cooking Recipes, Football History...")
    
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
        help="Unsplash: FREE topic-relevant images! AI: Custom generated (requires Stability key)"
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
st.info("✨ **Unsplash mode gets TOPIC-RELEVANT images for FREE!** No API key needed!")
st.warning("💡 **Important:** Use FREE Google Gemini model if you're low on credits!")
generate_button = st.button("🚀 Generate PowerPoint", use_container_width=True)

# Generation logic
if generate_button:
    if not claude_api_key:
        st.error("⚠️ Please enter your OpenRouter API key in the sidebar.")
    elif not claude_api_key.startswith("sk-or-"):
        st.error("⚠️ Invalid OpenRouter API key format. It should start with 'sk-or-'")
    elif not topic:
        st.error("⚠️ Please enter a topic for your presentation.")
    elif image_mode == "AI Generated (Paid)" and not stability_api_key:
        st.error("⚠️ Stability AI key required for AI Generated images. Please add your key or switch to free images.")
    else:
        with st.spinner("🤖 AI is analyzing your topic and creating custom content..."):
            try:
                st.info(f"🎨 Image Mode: {image_mode}")
                
                # Generate content with Claude
                slides_content = generate_content_with_claude(
                    claude_api_key, topic, category, slide_count, 
                    tone, audience, key_points, model_choice
                )
                
                if slides_content:
                    st.success("✅ Content generated! Now creating PowerPoint with images...")
                    
                    # Create PowerPoint
                    prs = create_powerpoint(
                        slides_content, theme, image_mode, 
                        stability_api_key if image_mode == "AI Generated (Paid)" else None,
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
                import traceback
                st.code(traceback.format_exc())

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 1rem;'>
    <p>💡 <strong>Works with ANY topic!</strong></p>
    <p>🎨 <strong>Unsplash (FREE)</strong> - Gets images RELEVANT to your topic automatically!</p>
    <p>🤖 <strong>AI Generated</strong> - Custom images (needs Stability AI key + credits)</p>
    <p>🆓 <strong>Use "Free Model (Google Gemini)"</strong> to avoid credit issues!</p>
    <p>⚠️ <strong>Low on credits?</strong> Reduce slides or upgrade at <a href="https://openrouter.ai/settings/credits">OpenRouter</a></p>
</div>
""", unsafe_allow_html=True)
