import streamlit as st

class Login:

    def __init__(self):
        pass

    def login(self):
        st.sidebar.title("🔐 Login")
        username = st.sidebar.text_input("Gebruikersnaam")
        password = st.sidebar.text_input("Wachtwoord", type="password")
        login_knop = st.sidebar.button("Inloggen")

        if login_knop:
            if username in st.secrets["users"] and st.secrets["users"][username] == password:
                st.session_state["logged_in"] = True
                st.session_state["user"] = username
                st.rerun()
            else:
                st.sidebar.error("Ongeldige gebruikersnaam of wachtwoord")
    
    def require_login():
        if "logged_in" not in st.session_state or not st.session_state["logged_in"]:
            login = Login()
            login.login()
            st.stop()