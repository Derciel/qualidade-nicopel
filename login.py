import streamlit as st
from urllib.parse import urlencode

st.set_page_config(page_title="Login", layout="centered")

st.image("https://cdn-icons-png.flaticon.com/512/732/732221.png", width=100)
st.title("Bem-vindo ao Dashboard de Qualidade")

client_id = st.secrets["oauth_microsoft"]["client_id"]
tenant_id = st.secrets["oauth_microsoft"]["tenant_id"]
redirect_uri = st.secrets["oauth_microsoft"]["redirect_uri"]

params = {
    "client_id": client_id,
    "response_type": "code",
    "redirect_uri": redirect_uri,
    "response_mode": "query",
    "scope": "User.Read openid profile email",
    "state": "12345"
}
auth_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/authorize?" + urlencode(params)

st.markdown(
    f"""
    <div style='text-align: center; margin-top: 30px;'>
        <a href="{auth_url}" style='
            background-color:#2F2F2F;
            color:white;
            padding:12px 24px;
            border-radius:8px;
            font-weight:bold;
            text-decoration:none;
            display:inline-block;
        '>
            Entrar com Microsoft
        </a>
    </div>
    """,
    unsafe_allow_html=True
)
