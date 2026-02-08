from flask import Flask, render_template, request, redirect
from pymongo import MongoClient

app = Flask(__name__)

# ---------------- MONGODB CONNECTION ----------------

client = MongoClient("mongodb://localhost:27017/")
db = client["contact_app"]
collection = db["contacts"]

# ---------------- ROUTES ----------------

@app.route('/')
def index():
    contacts = list(collection.find({}, {"_id": 0}))
    return render_template("index.html", contacts=contacts)


@app.route('/add', methods=['GET', 'POST'])
def add_contact():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        phone = request.form['phone']

        collection.insert_one({
            "name": name,
            "email": email,
            "phone": phone
        })

        return redirect('/')

    return render_template("add.html")


if __name__ == '__main__':
    app.run(debug=True)
