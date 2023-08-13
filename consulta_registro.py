#!/usr/bin/env python
# coding: utf-8

# # Employee Records
#

# A graphical application will be created using Python to manage an employee database. The management tasks include establishing a database connection, creating an employee table, inserting, querying, updating, and deleting employee records and information within the company.

# In[1]:


# Library Imports
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QTextCursor
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton, QTextEdit
import sys
import sqlite3
import pandas as pd
from sqlite3 import Error
from tkinter import *
from tkinter import messagebox
from PyQt5.QtWidgets import QFileDialog, QComboBox
from fpdf import FPDF
from docx import Document


# # Database Connection Creation

# In[2]:


def create_connection():
    conn = None
    try:
        conn = sqlite3.connect('employee_db.sqlite')
        print(sqlite3.version)
    except Error as e:
        print(e)
    return conn


# # Employee Table Creation

# In[3]:


def create_table(conn):
    try:
        sql = ''' CREATE TABLE IF NOT EXISTS employees (
                                        id integer PRIMARY KEY,
                                        name text NOT NULL,
                                        position text NOT NULL,
                                        salary real
                                    ); '''
        conn.cursor().execute(sql)
    except Error as e:
        print(e)


# ## Employee Insertion

# In[4]:


def insert_employee(conn, employee):
    sql = ''' INSERT INTO employees(name, position, salary)
              VALUES(?,?,?) '''
    cur = conn.cursor()
    cur.execute(sql, employee)
    conn.commit()
    return cur.lastrowid


# # Employee Query

# In[5]:


def query_employees(conn):
    cur = conn.cursor()
    cur.execute("SELECT * FROM employees")
    rows = cur.fetchall()
    for row in rows:
        print(row)


# # Employee Update

# In[6]:


def update_employee(conn, employee):
    sql = '''
    UPDATE employees
    SET name = ?,
        position = ?,
        salary = ?
    WHERE id = ?
    '''
    cur = conn.cursor()
    cur.execute(sql, employee)
    conn.commit()


# # Employee Deletion

# In[7]:


def delete_employee(conn, id):
    sql = 'DELETE FROM employees WHERE id = ?'
    cur = conn.cursor()
    cur.execute(sql, (id,))
    conn.commit()


# # Application Window Creation

# In[8]:


def create_window():
    window = Tk()
    window.title("Employee Management")
    window.geometry("800x600")
    return window


# # User Interface Components

# In[9]:


# Function to create a database connection

def create_connection():
    conn = None
    try:
        conn = sqlite3.connect('empleados_db.sqlite')
        print(sqlite3.version)
    except Error as e:
        print(e)
    return conn

# Function to create the employees table


def create_table(conn):
    try:
        sql = '''CREATE TABLE IF NOT EXISTS empleados (
                    id integer PRIMARY KEY,
                    nombre text NOT NULL,
                    cargo text NOT NULL,
                    salario real
                );'''
        conn.cursor().execute(sql)
    except Error as e:
        print(e)

# Function to insert an employee into the database


def insert_employee(conn, employee):
    sql = '''INSERT INTO empleados(nombre, cargo, salario)
            VALUES(?,?,?)'''
    cur = conn.cursor()
    cur.execute(sql, employee)
    conn.commit()
    return cur.lastrowid

# Function to retrieve all employees from the database


def retrieve_employees(conn):
    cur = conn.cursor()
    cur.execute("SELECT * FROM empleados")
    rows = cur.fetchall()
    for row in rows:
        print(row)

# Function to update an employee in the database


def update_employee(conn, employee):
    sql = '''
    UPDATE empleados
    SET nombre = ?,
        cargo = ?,
        salario = ?
    WHERE id = ?
    '''
    cur = conn.cursor()
    cur.execute(sql, employee)
    conn.commit()

# Function to delete an employee from the database


def delete_employee(conn, id):
    sql = 'DELETE FROM empleados WHERE id = ?'
    cur = conn.cursor()
    cur.execute(sql, (id,))
    conn.commit()

# Class for the main application window


class MainWindow(QMainWindow):
    def __init__(self, conn):
        super().__init__()
        self.conn = conn
        self.setWindowTitle("Employee Management")
        self.setGeometry(100, 100, 800, 600)
        self.setStyleSheet("background-color: #20232a; color: #61dafb;")

        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)

        self.name_text = QLineEdit()
        self.position_text = QLineEdit()
        self.salary_text = QLineEdit()
        self.results_text = QTextEdit()
        self.id_text = QLineEdit()
        self.export_format_combo = QComboBox()
        self.export_format_combo.addItems(["PDF", "Excel", "Word"])

        main_layout.addWidget(QLabel("Employee Name:"))
        main_layout.addWidget(self.name_text)

        main_layout.addWidget(QLabel("Employee Position:"))
        main_layout.addWidget(self.position_text)

        main_layout.addWidget(QLabel("Employee Salary:"))
        main_layout.addWidget(self.salary_text)

        main_layout.addWidget(QPushButton(
            "Insert Employee", clicked=self.insert_employee))

        main_layout.addWidget(QLabel("Query Results:"))
        main_layout.addWidget(self.results_text)

        main_layout.addWidget(QPushButton(
            "Retrieve Employees", clicked=self.retrieve_employees))

        main_layout.addWidget(QLabel("Employee ID for Update/Delete:"))
        main_layout.addWidget(self.id_text)

        main_layout.addWidget(QPushButton(
            "Load Employee", clicked=self.load_employee))
        main_layout.addWidget(QPushButton(
            "Update Employee", clicked=self.update_employee))
        main_layout.addWidget(QPushButton(
            "Delete Employee", clicked=self.delete_employee))

        main_layout.addWidget(self.export_format_combo)
        main_layout.addWidget(QPushButton("Export", clicked=self.export_data))

        self.setCentralWidget(main_widget)

    def export_data(self):
        export_format = self.export_format_combo.currentText()

        if export_format == "Word":
            self.export_to_word()
        elif export_format == "PDF":
            self.export_to_pdf()
        elif export_format == "Excel":
            self.export_to_excel()
        

    def export_to_pdf(self):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM empleados")
        rows = cur.fetchall()

        pdf_filename, _ = QFileDialog.getSaveFileName(
            None, "Save PDF File", "", "PDF Files (*.pdf)")

        if pdf_filename:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            for row in rows:
                pdf.cell(
                    200, 10, f"ID: {row[0]}, Nombre: {row[1]}, Cargo: {row[2]}, Salario: {row[3]}", ln=True)

            pdf.output(pdf_filename)

            QMessageBox.information(self, "Information",
                                    "Data exported to PDF successfully.")


    def export_to_excel(self):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM empleados")
        rows = cur.fetchall()
        df = pd.DataFrame(rows, columns=["ID", "Nombre", "Cargo", "Salario"])

        excel_filename, _ = QFileDialog.getSaveFileName(
            None, "Save Excel File", "", "Excel Files (*.xlsx)")

        if excel_filename:
            try:
                df.to_excel(excel_filename, index=False)
                QMessageBox.information(
                    self, "Information", "Data exported to Excel successfully.")
            except Exception as e:
                QMessageBox.warning(
                    self, "Error", f"Error exporting data to Excel:\n{str(e)}")

    def export_to_word(self):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM empleados")
        rows = cur.fetchall()

        docx_filename, _ = QFileDialog.getSaveFileName(
            None, "Save Word File", "", "Word Files (*.docx)")

        if docx_filename:
            doc = Document()

            for row in rows:
                doc.add_paragraph(
                    f"ID: {row[0]}, Nombre: {row[1]}, Cargo: {row[2]}, Salario: {row[3]}")

            doc.save(docx_filename)

            QMessageBox.information(self, "Information",
                                    "Data exported to Word successfully.")


    def insert_employee(self):
        name = self.name_text.text()
        position = self.position_text.text()
        salary = float(self.salary_text.text())
        employee = (name, position, salary)
        employee_id = insert_employee(self.conn, employee)
        QMessageBox.information(self, "Information",
                                f"Employee inserted with ID: {employee_id}")
        self.name_text.clear()
        self.position_text.clear()
        self.salary_text.clear()

    def retrieve_employees(self):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM empleados")
        rows = cur.fetchall()
        self.results_text.clear()
        for row in rows:
            self.results_text.insertPlainText(str(row) + '\n')

    def load_employee(self):
        employee_id = int(self.id_text.text())
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM empleados WHERE id=?", (employee_id,))
        row = cur.fetchone()
        if row is not None:
            self.name_text.setText(row[1])
            self.position_text.setText(row[2])
            self.salary_text.setText(str(row[3]))
        else:
            QMessageBox.warning(
                self, "Error", "No employee found with that ID")

    def update_employee(self):
        employee_id = int(self.id_text.text())
        response = QMessageBox.question(
            self, "Confirm Update", "Are you sure you want to update this employee?")
        if response == QMessageBox.Yes:
            name = self.name_text.text()
            position = self.position_text.text()
            salary = float(self.salary_text.text())
            employee = (name, position, salary, employee_id)
            update_employee(self.conn, employee)
            QMessageBox.information(self, "Information", "Employee updated.")
            self.name_text.clear()
            self.position_text.clear()
            self.salary_text.clear()

    def delete_employee(self):
        employee_id = int(self.id_text.text())
        response = QMessageBox.question(
            self, "Confirm Deletion", "Are you sure you want to delete this employee?")
        if response == QMessageBox.Yes:
            delete_employee(self.conn, employee_id)
            QMessageBox.information(self, "Information", "Employee deleted.")
            self.name_text.clear()
            self.position_text.clear()
            self.salary_text.clear()


if __name__ == '__main__':
    conn = create_connection()
    create_table(conn)

    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # Apply the Fusion style
    window = MainWindow(conn)
    window.show()

    sys.exit(app.exec_())


# It utilizes the PyQt5 library for creating the application window, handling user input, and displaying information.
#
# Please note that the code assumes the existence of a SQLite database file named "employee_db.sqlite" in the same directory as the application file.
#
# Feel free to explore the code and modify it according to your needs.
