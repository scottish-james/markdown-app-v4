import streamlit as st
def setup_enhanced_sidebar():
    """Enhanced sidebar setup with Claude Sonnet 4 as the primary AI provider."""
    with st.sidebar:
        st.header("Document to Markdown")
        st.write("AI-Powered Document Conversion")

        # Supported formats in an expander
        with st.expander("Supported Formats"):
            from src.ui.components import display_supported_formats
            display_supported_formats()

        # Enhancement options
        st.subheader("AI Enhancement")
        enhance_markdown = st.checkbox("Enhance with Claude Sonnet 4", value=True,
                                       help="Use Claude to improve markdown formatting and structure")

        api_key = None
        if enhance_markdown:
            api_key = st.text_input(
                "Anthropic API Key",
                type="password",
                help="Enter your Anthropic API key for Claude enhancement"
            )

            # Clear instructions for getting API key
            if not api_key:
                st.info("💡 **Need an API key?**")
                st.markdown("""
                1. Visit [console.anthropic.com](https://console.anthropic.com/)
                2. Create an account or sign in
                3. Generate a new API key
                4. Paste it above to unlock AI enhancement
                """)
            else:
                st.success("✅ Claude Sonnet 4 ready!")

        # Developer info
        st.sidebar.markdown("---")
        st.sidebar.markdown("""
        **Developed by:** James Taylor  
        **Powered by:** Claude Sonnet 4  
        """)

    return enhance_markdown, api_key