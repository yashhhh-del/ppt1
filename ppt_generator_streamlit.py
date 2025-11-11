import streamlit as st
import anthropic
import os

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
st.markdown('<div class="main-header">📊 AI PowerPoint Generator</div>', unsafe_allow_html=True)
st.markdown("---")

# Sidebar for API Key
with st.sidebar:
    st.header("⚙️ Configuration")
    api_key = st.text_input("Anthropic API Key", type="password", help="Enter your Anthropic API key")
    st.markdown("---")
    st.markdown("### 📖 How to Use")
    st.markdown("""
    1. Enter your Anthropic API key
    2. Fill in presentation details
    3. Click 'Generate Presentation'
    4. Copy the structured output
    """)
    st.markdown("---")
    st.markdown("Made with ❤️ using Claude AI")

# Main content
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📝 Presentation Details")
    
    topic = st.text_input("Topic *", placeholder="e.g., Digital Marketing Strategy 2025")
    
    category = st.selectbox(
        "Category *",
        ["Business", "Pitch", "Marketing", "Technical", "Academic", "Training"]
    )
    
    slide_count = st.number_input("Number of Slides *", min_value=3, max_value=30, value=10)
    
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
        ["Stock", "AI", "Mixed", "None"]
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
generate_button = st.button("🚀 Generate Presentation", use_container_width=True)

# Generation logic
if generate_button:
    if not api_key:
        st.error("⚠️ Please enter your Anthropic API key in the sidebar.")
    elif not topic:
        st.error("⚠️ Please enter a topic for your presentation.")
    else:
        with st.spinner("🎨 Generating your professional presentation..."):
            try:
                # Initialize Anthropic client
                client = anthropic.Anthropic(api_key=api_key)
                
                # Prepare the prompt
                prompt = f"""You are an expert corporate PPT creator.
Your task is to generate a complete and professional PowerPoint presentation in structured form.
------------------------------------------------
INPUT
Topic: {topic}
Category: {category}
Slide Count: {slide_count}
Tone: {tone}
Audience: {audience}
Theme Style: {theme}
Image Mode: {image_mode}
English Variant: {english_variant}
Additional Key Points (optional):
{key_points if key_points else "None"}
------------------------------------------------
OUTPUT FORMAT (REQUIRED)
Return presentation in this structure:
Slide 1:
Title: <max 8 words>
Bullets:
(No bullets on cover slide)
Image Suggestion: <Describe visual style, NO TEXT in image>

Slide 2:
Title:
Bullets:
- <max 12 words each>
- 
- 
Image Suggestion: <Abstract / Stock Visual / or None>

Continue until slide count completed.
------------------------------------------------
RULES
1. All bullets must be short, clear, and professional (≤ 12 words).
2. Maintain consistent tone and message flow.
3. Grammar must be 100% correct.
4. If Image Mode = None → Skip "Image Suggestion".
5. If Image Mode = AI → Describe abstract clean visuals (NO faces, NO logos, NO text).
6. If Image Mode = Stock → Suggest real-world images relevant to context.
7. Final slide should be a conclusion or call-to-action.
------------------------------------------------
Now generate the PPT structure."""

                # Call Claude API
                message = client.messages.create(
                    model="claude-sonnet-4-20250514",
                    max_tokens=8000,
                    messages=[
                        {"role": "user", "content": prompt}
                    ]
                )
                
                # Extract response
                response_text = message.content[0].text
                
                # Display results
                st.success("✅ Presentation generated successfully!")
                st.markdown("---")
                
                st.markdown('<div class="output-box">', unsafe_allow_html=True)
                st.markdown("### 📄 Your Presentation Structure")
                st.markdown(response_text)
                st.markdown('</div>', unsafe_allow_html=True)
                
                # Download button
                st.download_button(
                    label="📥 Download as Text File",
                    data=response_text,
                    file_name=f"{topic.replace(' ', '_')}_presentation.txt",
                    mime="text/plain"
                )
                
            except anthropic.AuthenticationError:
                st.error("❌ Authentication failed. Please check your API key.")
            except anthropic.APIError as e:
                st.error(f"❌ API Error: {str(e)}")
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 1rem;'>
    <p>💡 <strong>Tip:</strong> For best results, provide specific topics and clear key points.</p>
    <p>Get your Anthropic API key from <a href='https://console.anthropic.com/' target='_blank'>console.anthropic.com</a></p>
</div>
""", unsafe_allow_html=True)