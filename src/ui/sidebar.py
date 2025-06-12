import streamlit as st


def setup_enhanced_sidebar():
    """Enhanced sidebar setup with Claude Sonnet 4 as the primary AI provider."""
    with st.sidebar:
        # st.header("Document to Markdown")
        # st.write("AI-Powered Document Conversion")

        # Enhancement options
        st.subheader("AI Enhancement")
        enhance_markdown = st.toggle("Enhance with Claude Sonnet 4", value=True,
                                     help="Use Claude to improve markdown formatting and structure")

        api_key_claude = None
        if enhance_markdown:
            api_key_claude = st.text_input(
                "Anthropic API Key",
                type="password",
                help="Enter your Anthropic API key for Claude enhancement"
            )

            # Clear instructions for getting API key
            if not api_key_claude:
                st.info("💡 **Need an API key?**")
                st.markdown("""
                1. Visit [console.anthropic.com](https://console.anthropic.com/)
                2. Create an account or sign in
                3. Generate a new API key
                4. Paste it above to unlock AI enhancement
                """)
            else:
                st.success("✅ Claude Sonnet 4 ready!")

        enhance_diagram = st.toggle("AI Diagrams using OpenAI", value=True,
                                    help="Use OpenAI to create diagrams")

        api_key_openai = None

        if enhance_diagram:
            api_key_openai = st.text_input(
                "OpenAI API Key [not working yet]",
                type="password",
                help="Enter your OpenAI API key for Diagram creation"
            )

            # Clear instructions for getting API key
            if not api_key_openai:
                st.info("💡 **Need an API key?**")
                st.markdown("""
                1. Visit [OpenAI platform](https://platform.openai.com/)
                2. Sign in and go to **API Keys**
                3. Click **Create new secret key**
                4. Copy and save it securely
                """)
            else:
                st.success("✅ OpenAI ready!")

        # Developer info
        st.sidebar.markdown("---")
        st.sidebar.markdown("""
        **Developed by:** James Taylor  
        **Powered by:** Claude Sonnet 4  
        """)

    return enhance_markdown, api_key_claude
