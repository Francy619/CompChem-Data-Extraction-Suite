import streamlit as st

# Set a minimal page config
st.set_page_config(page_title="Redirecting...", page_icon="🔄")

# Define the new target URL
new_url = "https://ctcc-sdck.streamlit.app/"

# Inject HTML and JavaScript to force an immediate redirect
redirect_code = f"""
<meta http-equiv="refresh" content="0; url={new_url}">
<script>
    window.location.replace("{new_url}");
</script>
"""
st.markdown(redirect_code, unsafe_allow_html=True)

# Provide a clickable fallback link in case the user's browser blocks JS redirects
st.warning("🔄 **We have moved!** You are being redirected to our new URL...")
st.markdown(f"If you are not redirected automatically, **[click here to go to the new CompChem Suite]({new_url})**.")
