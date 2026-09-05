from werkzeug.security import generate_password_hash, check_password_hash
from flask import Flask, render_template, request, redirect
import os
from werkzeug.utils import secure_filename
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)

app.config["UPLOAD_FOLDER"] = os.path.join(app.root_path, "static", "uploads")

# Database configuration
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///stockflow.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

# Users table
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
class Product(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    category = db.Column(db.String(50), nullable=False)
    price = db.Column(db.Float, nullable=False)
    quantity = db.Column(db.Integer, nullable=False) 
    supplier_id = db.Column(db.Integer, db.ForeignKey("supplier.id")) 
    image = db.Column(db.String(200))
class Supplier(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    company = db.Column(db.String(100), nullable=False)
    phone = db.Column(db.String(15), nullable=False)      

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":
        email = request.form["email"]
        password = request.form["password"]

        user = User.query.filter_by(email=email).first()

        if user and check_password_hash(user.password, password):
            return redirect("/dashboard")
        else:
            return "Invalid Email or Password"

    return render_template("login.html")

from sqlalchemy import func
import json
from openpyxl import Workbook
from flask import send_file
def indian_currency(num):
    num = int(num)
    s = str(num)

    if len(s) <= 3:
        return s

    last3 = s[-3:]
    rest = s[:-3]

    parts = []
    while len(rest) > 2:
        parts.insert(0, rest[-2:])
        rest = rest[:-2]

    if rest:
        parts.insert(0, rest)

    return ",".join(parts + [last3])

@app.route("/dashboard")
def dashboard():

    total_products = Product.query.count()

    low_stock = Product.query.filter(Product.quantity < 10).count()

    total_suppliers = Supplier.query.count()

    inventory_value = db.session.query(
        func.sum(Product.price * Product.quantity)
    ).scalar() or 0

    # Bar chart data
    category_data = db.session.query(
        Product.category,
        func.count(Product.id)
    ).group_by(Product.category).all()

    labels = [item[0] for item in category_data]
    values = [item[1] for item in category_data]

    # Recent products
    recent_products = Product.query.order_by(
        Product.id.desc()
    ).limit(5).all()

    return render_template(
        "dashboard.html",
        total_products=total_products,
        low_stock=low_stock,
        total_suppliers=total_suppliers,
        inventory_value=indian_currency(inventory_value),
        chart_labels=json.dumps(labels),
        chart_values=json.dumps(values),
        recent_products=recent_products
    )

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        user = User(
            name=request.form["name"],
            email=request.form["email"],
            password=generate_password_hash(request.form["password"])
        )

        db.session.add(user)
        db.session.commit()

        return redirect("/login")

    return render_template("register.html")
@app.route("/add-product", methods=["GET", "POST"])
def add_product():

    suppliers = Supplier.query.all()

    if request.method == "POST":

        image = request.files["image"]
        filename = ""

        if image and image.filename:
            filename = secure_filename(image.filename)
            image.save(os.path.join(app.config["UPLOAD_FOLDER"], filename))

        product = Product(
            name=request.form["name"],
            category=request.form["category"],
            price=float(request.form["price"]),
            quantity=int(request.form["quantity"]),
            supplier_id=int(request.form["supplier"]),
            image=filename
        )

        db.session.add(product)
        db.session.commit()

        return redirect("/products")

    return render_template(
        "add_product.html",
        suppliers=suppliers
    )
from sqlalchemy import or_

@app.route("/products")
def products():

    search = request.args.get("search", "")

    if search:
        all_products = Product.query.filter(
            or_(
                Product.name.ilike(f"%{search}%"),
                Product.category.ilike(f"%{search}%")
            )
        ).all()
    else:
        all_products = Product.query.all()

    return render_template(
        "products.html",
        products=all_products,
        search=search
    )

@app.route("/delete-product/<int:id>")
def delete_product(id):
    product = Product.query.get_or_404(id)

    db.session.delete(product)
    db.session.commit()

    return redirect("/products")
@app.route("/edit-product/<int:id>", methods=["GET", "POST"])
def edit_product(id):

    product = Product.query.get_or_404(id)

    if request.method == "POST":
        product.name = request.form["name"]
        product.category = request.form["category"]
        product.price = float(request.form["price"])
        product.quantity = int(request.form["quantity"])

        db.session.commit()

        return redirect("/products")

    return render_template("edit_product.html", product=product)
@app.route("/add-supplier", methods=["GET", "POST"])
def add_supplier():

    if request.method == "POST":

        supplier = Supplier(
            name=request.form["name"],
            company=request.form["company"],
            phone=request.form["phone"]
        )

        db.session.add(supplier)
        db.session.commit()

        return redirect("/suppliers")

    return render_template("add_supplier.html")
@app.route("/suppliers")
def suppliers():

    all_suppliers = Supplier.query.all()

    return render_template(
        "suppliers.html",
        suppliers=all_suppliers
    )
@app.route("/export")
def export_excel():

    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    # Header
    ws.append([
        "ID",
        "Product",
        "Category",
        "Price",
        "Quantity"
    ])

    # Data
    products = Product.query.all()

    for p in products:
        ws.append([
            p.id,
            p.name,
            p.category,
            p.price,
            p.quantity
        ])

    file_name = "StockFlow_Inventory.xlsx"
    wb.save(file_name)

    return send_file(file_name, as_attachment=True)

if __name__ == "__main__":
    with app.app_context():
        db.create_all()   # Creates stockflow.db automatically
    app.run(debug=True)