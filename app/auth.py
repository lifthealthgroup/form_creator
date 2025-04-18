from flask import Blueprint, request, redirect, url_for, session, render_template
from functools import wraps
import os

PASSWORD = os.getenv("FORM_CREATOR_PASSWORD")

auth = Blueprint("auth", __name__)

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("auth.login"))
        return f(*args, **kwargs)
    return decorated_function

@auth.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        if request.form["password"] == PASSWORD:
            session["logged_in"] = True
            return redirect(url_for("index"))
        else:
            error = "Incorrect password"
    return render_template("login.html", error=error)