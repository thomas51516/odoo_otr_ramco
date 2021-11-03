# -*- coding: utf-8 -*-
from flask import Flask, render_template, request
from flask_sqlalchemy import SQLAlchemy
from flask_odoo import Odoo
from flask_mail import Mail, Message
import os
from datetime import datetime
import paramiko

app = Flask(__name__)

app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///db.sqlite"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

app.config["ODOO_URL"] = "http://ramcootr.odoo.com"
app.config["ODOO_DB"] = "edi228-ramco-test2-3477732"
app.config["ODOO_USERNAME"] = "odootogo@gmail.com"
app.config["ODOO_PASSWORD"] = "odootogo@gmail.com"

# app.config["ODOO_URL"] = "http://10.20.20.214:8069"
# app.config["ODOO_DB"] = "Demo"
# app.config["ODOO_USERNAME"] = "Roots"
# app.config["ODOO_PASSWORD"] = "Roots"
odoo = Odoo(app)

app.config['MAIL_SERVER']='smtp.gmail.com'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = 'admin@ramco.tg'
app.config['MAIL_PASSWORD'] = 'vcudkpaglxtqhalu'
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
app.config['MAIL_ASCII_ATTACHMENTS'] = True
mail = Mail(app)

class OdooConfig(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    slug =  db.Column(db.String, default="otr",nullable=False)
    server_name = db.Column(db.String, nullable=False)
    host_name = db.Column(db.String, nullable=False)
    username = db.Column(db.String, unique=True, nullable=False)
    password = db.Column(db.String, nullable=False)


@app.route('/', methods=['GET','POST'])
def home():
    if request.method == "POST":
        config = OdooConfig()
        config.host_name = request.form['host_name']
        config.server_name = request.form['server_name']
        config.username = request.form['username']
        config.password = request.form['odoopass']
        db.session.add(config)
        db.session.commit()
    servers = OdooConfig.query.all()
    return render_template('index.html', servers=servers)

class PosOrder(odoo.Model):
    _name = "pos.order"
    _domain = []

    name = odoo.StringType()
    date_order = odoo.DateTimeType()
    user_id = odoo.Many2oneType()
    partner_id = odoo.Many2oneType()
    lines = odoo.One2manyType()
    state = odoo.StringType()

class SaleOrder(odoo.Model):
    _name = "sale.order"
    _domain = []

    name = odoo.StringType()
    date_order = odoo.DateTimeType()
    user_id = odoo.Many2oneType()
    partner_id = odoo.Many2oneType()
    order_line = odoo.One2manyType()
    state = odoo.StringType()

class PosOrderLine(odoo.Model):
    _name = "pos.order.line"
    _domain = []

    product_id = odoo.Many2oneType()
    price_unit = odoo.FloatType()
    qty = odoo.FloatType()

class SaleOrderLine(odoo.Model):
    _name = "sale.order.line"
    _domain = []

    product_id = odoo.Many2oneType()
    price_unit = odoo.FloatType()
    product_uom_qty = odoo.FloatType()

class Product(odoo.Model):
    _name = "product.product"
    _domain = []
    name = odoo.StringType()
    default_code = odoo.StringType()
    display_name = odoo.StringType()
    lst_price = odoo.FloatType()
    categ_id = odoo.Many2oneType()

class ProductCateg(odoo.Model):
    _name = "product.category"
    _domain = []
    name = odoo.StringType()


class Parter(odoo.Model):
    _name = "res.partner"
    _domain = []
    name = odoo.StringType()
    phone = odoo.StringType()
    flight_number = odoo.StringType()
    flight_date = odoo.DateType()
    traveler_name = odoo.StringType()
    traveler_first_name = odoo.StringType()
    departure_place = odoo.StringType()
    destination_place = odoo.StringType()
    type_id = odoo.StringType()
    id_number = odoo.StringType()
    nationality = odoo.StringType()

def generate_xml_file(order, order_lines, product, partner):
    filepath ="static/" + order.name.replace('/','-') + '_' + order.date_order.strftime('%m/%d/%Y').replace('/','-') + ".xml"
    if os.path.exists(filepath):
        os.remove(filepath)
        open(filepath, "x")

    f = open(filepath, "w")
    print(partner.flight_number)
    body = """<?xml version="1.0" encoding="utf-8"?>\n"""
    body +="<DFS_SYDONIA>\n"
    body +="\t<GENERAL_SEGMENT>\n"
    body +="\t\t<reference>"+ order.name + "</reference>\n"
    body +="\t\t<ref_dat>" + order.date_order.strftime('%m/%d/%Y') +"</ref_dat>\n"
    body +="\t\t<type_action>" + "INSERT" +"</type_action>\n"
    body +="\t\t<statut>" + order.state + "</statut>\n"
    body +="\t\t<date_envoie>" + datetime.today().strftime('%m/%d/%Y') + "</date_envoie>\n"
    body +="\t\t<code_bureau>" + 'TG122' + "</code_bureau>\n"
    body +="\t\t<libelle_bureau>" + "BUREAU DE L'AEROPORT" + "</libelle_bureau>\n"
    body +="\t\t<code_importateur>" + '1000166599' + "</code_importateur>\n"
    body +="\t\t<nom_importateur> RAMCO </nom_importateur>\n"
    body +="\t</GENERAL_SEGMENT>\n\t<VENTES>\n"
    body +="\t\t<numero_vol>" + partner.flight_number + "</numero_vol>\n"
    body +="\t\t<date_vol>" + order.date_order.strftime('%m/%d/%Y') + "</date_vol>\n"
    body +="\t\t<nom_voyageur>" + partner.traveler_name + "</nom_voyageur>\n"
    body +="\t\t<prenom_voyageur>" + partner.traveler_first_name + "</prenom_voyageur>\n"
    body +="\t\t<lieu_depart>" + partner.departure_place + "</lieu_depart>\n"
    body +="\t\t<lieu_destination>" + partner.destination_place + "</lieu_destination>\n"
    body +="\t\t<type_piece>" + str(partner.type_id) + "</type_piece>\n"
    body +="\t\t<numero_piece>" + partner.id_number + "</numero_piece>\n"
    body +="\t\t<nationalite>" + partner.nationality + "</nationalite>\n"
    for line in order_lines:
        qty = 0
        if(isinstance(order, PosOrder)):
            qty = line.qty
        else:
            qty = line.product_uom_qty
        product = Product.search_by_id(line.product_id[0])
        body += "\t\t<ARTICLE>\n\t\t\t<nomenclature>" + product.default_code + "</nomenclature>\n"
        body += "\t\t\t<desc_article>" +  product.name + "</desc_article>\n"
        body += "\t\t\t<categorie_article>" +  product.categ_id[1] + "</categorie_article>\n"
        body += "\t\t\t<quantite>" + str(qty) + "</quantite>\n"
        body += "\t\t\t<montant>" + str(line.price_unit) + "</montant>\n"
        body += "\t\t\t<date_vente>" + order.date_order.strftime('%m/%d/%Y') + "</date_vente>\n"
        body += "\t\t</ARTICLE>\n"
    body += "\t</VENTES>\n"
    body +="</DFS_SYDONIA>"
    f.write(body)
    f.close()

@app.route('/otr/<int:id>')
def sendfile(id):
    order = PosOrder.search_by_id(id)
    order_lines = PosOrderLine.search_read([["order_id", "=", order.id]])
    product = order_lines[0].product_id[1]
    myid = 1
    if order.partner_id:
        myid = order.partner_id[0]
    partner = Parter.search_by_id(myid)
    generate_xml_file(order, order_lines, product, partner)
    filename = order.name.replace('/','-') + '_' + order.date_order.strftime('%m/%d/%Y').replace('/','-') + '.xml'
    sendToSFtp(filename)
    return render_template('odoo.html',  partner = partner, order=order, date = datetime.today().strftime('%Y-%m-%d-%H:%M:%S'))


@app.route('/otr/sale/<int:id>')
def sendSalefile(id):
    order = SaleOrder.search_by_id(id)
    order_lines = SaleOrderLine.search_read([["order_id", "=", order.id]])
    product = order_lines[0].product_id[1]
    myid = 1
    if order.partner_id:
        myid = order.partner_id[0]
    partner = Parter.search_by_id(myid)
    generate_xml_file(order, order_lines, product, partner)
    sendToSFtp(order)
    return render_template('odoo.html', partner = partner, order=order, date = datetime.today().strftime('%Y-%m-%d-%H:%M:%S'))


def sendToSFtp(file_name):
    filepath ="static/" + file_name
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(hostname='41.207.181.214',username='ramco', password='PwG.2Kk:r648Hx', port=8202, allow_agent=False, look_for_keys=False)
    sftp = client.open_sftp()
    sftp.put(filepath, file_name)

if __name__=='__main__':
    db.create_all()
    app.run(debug=True)