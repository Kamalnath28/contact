from flask import Flask, render_template, request, redirect, url_for, send_file
from pymongo import MongoClient
from bson.objectid import ObjectId
import re
import openpyxl
from io import BytesIO

app = Flask(__name__)

# MongoDB Connection
client = MongoClient("mongodb://localhost:27017/")
db = client["contact_app"]
collection = db["contacts"]


# ---------------- CONTACTS PAGE (SEARCH + FILTER) ----------------
@app.route('/')
def index():
    search_query = request.args.get('search')
    gender_filter = request.args.get('gender')
    city_filter = request.args.get('city')

    query = {}

    # Search logic
    if search_query:
        query["$or"] = [
            {"name": {"$regex": search_query, "$options": "i"}},
            {"email": {"$regex": search_query, "$options": "i"}},
            {"phone": {"$regex": search_query, "$options": "i"}}
        ]

    # Gender filter
    if gender_filter and gender_filter != "All":
        query["gender"] = gender_filter

    # City filter
    if city_filter and city_filter != "All":
        query["city"] = city_filter

    contacts = list(collection.find(query))

    return render_template("index.html", contacts=contacts)


# ---------------- ADD ----------------
@app.route('/add', methods=['GET', 'POST'])
def add_contact():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        phone = request.form['phone']
        gender = request.form['gender']
        city = request.form['city']

        # Validation
        if not re.match("^[A-Za-z ]+$", name):
            return "Invalid Name"

        if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
            return "Invalid Email"

        if not re.match("^[0-9]{10}$", phone):
            return "Phone must be 10 digits"

        collection.insert_one({
            "name": name,
            "email": email,
            "phone": phone,
            "gender": gender,
            "city": city
        })

        return redirect(url_for('index'))

    return render_template("add.html")


# ---------------- DELETE ----------------
@app.route('/delete/<id>')
def delete_contact(id):
    collection.delete_one({"_id": ObjectId(id)})
    return redirect(url_for('index'))


# ---------------- EDIT ----------------
@app.route('/edit/<id>', methods=['GET', 'POST'])
def edit_contact(id):
    contact = collection.find_one({"_id": ObjectId(id)})

    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        phone = request.form['phone']
        gender = request.form['gender']
        city = request.form['city']

        if not re.match("^[A-Za-z ]+$", name):
            return "Invalid Name"

        if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
            return "Invalid Email"

        if not re.match("^[0-9]{10}$", phone):
            return "Phone must be 10 digits"

        collection.update_one(
            {"_id": ObjectId(id)},
            {"$set": {
                "name": name,
                "email": email,
                "phone": phone,
                "gender": gender,
                "city": city
            }}
        )

        return redirect(url_for('index'))

    return render_template("edit.html", contact=contact)


# ---------------- EXPORT TO EXCEL ----------------
@app.route('/export')
def export_excel():
    contacts = list(collection.find())

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Contacts"

    # Updated header
    sheet.append(["Name", "Email", "Phone", "Gender", "City"])

    for contact in contacts:
        sheet.append([
            contact.get("name"),
            contact.get("email"),
            contact.get("phone"),
            contact.get("gender"),
            contact.get("city")
        ])

    file_stream = BytesIO()
    workbook.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name="contacts.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == '__main__':
    app.run(debug=True)
