import pyodbc
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import csv

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Set a secret key for flashing messages

# Database connection parameters
server = 'APL69954'
database = 'hardware'
username = 'apposcr'
password = '2#06A9a'
driver = '{SQL Server}'

# Create a connection to SQL Server
conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')

# Create a cursor to execute SQL commands
 # Programmer: APL69954;
cursor = conn.cursor()


@app.route('/')
def index():
    # Query the database to get the count of records in the HARDWARE_DATA table for each column
    cursor.execute('SELECT COUNT(SHOP_ID) AS COUNT_OF_SITES,COUNT(DVR) AS COUNT_OF_DVR,SUM(SYSTEM) AS SUM_OF_SYSTEM,SUM(PRINTER) AS SUM_OF_PRINTER,SUM(SCANNER) AS SUM_OF_SCANNER,SUM(DOME_CAMERA) AS SUM_OF_DOME_CAMERA,SUM(BULLET_CAMERA) AS SUM_OF_BULLET_CAMERA,SUM(SYSTEM + PRINTER + SCANNER) AS TOTAL_NO_OF_SPS,sum(DOME_CAMERA+BULLET_CAMERA) AS TOTAL_NO_OF_CAMERA FROM HARDWARE_DATA;')
    counts = cursor.fetchone()

    # Pass the counts to the HTML template
    return render_template('index.html', COUNT_OF_SITES=counts.COUNT_OF_SITES, COUNT_OF_DVR=counts.COUNT_OF_DVR, SUM_OF_SYSTEM=counts.SUM_OF_SYSTEM, SUM_OF_PRINTER=counts.SUM_OF_PRINTER, SUM_OF_SCANNER=counts.SUM_OF_SCANNER, SUM_OF_DOME_CAMERA=counts.SUM_OF_DOME_CAMERA, SUM_OF_BULLET_CAMERA=counts.SUM_OF_BULLET_CAMERA, TOTAL_NO_OF_SPS=counts.TOTAL_NO_OF_SPS, TOTAL_NO_OF_CAMERA=counts.TOTAL_NO_OF_CAMERA)


@app.route('/update', methods=['POST'])
def update_data():
    if request.method == 'POST':
        shopid = request.form['SHOP_ID']
        system = request.form['SYSTEM']
        printer = request.form['PRINTER']
        scanner = request.form['SCANNER']
        dome_camera = request.form['DOME_CAMERA']
        bullet_camera = request.form['BULLET_CAMERA']
        dvr = request.form['DVR']
        
        # Update the items for the specified SHOPID
        cursor.execute('''
            UPDATE HARDWARE_DATA
            SET SYSTEM = ? , PRINTER = ? , SCANNER = ? , DOME_CAMERA = ? , BULLET_CAMERA = ? ,DVR = ?
            WHERE SHOP_ID = ?
        ''', (system, printer, scanner, dome_camera,bullet_camera, dvr ,shopid))
        
        # Check the number of rows affected
        rows_affected = cursor.rowcount

        if rows_affected > 0:
            flash('UPDATED SUCCESSFULLY!', 'success')
        else:
            flash(f'UPDATE ERROR: SHOP_ID {shopid} NOT FOUND', 'error')

        conn.commit()
        return redirect(url_for('index'))

    
@app.route('/insert', methods=['POST'])
def insert_data():
    if request.method == 'POST':
        shopid = request.form['SHOP_ID']
        shopname = request.form['SHOP_NAME']
        system = request.form['SYSTEM']
        printer = request.form['PRINTER']
        scanner = request.form['SCANNER']
        dome_camera = request.form['DOME_CAMERA']
        bullet_camera = request.form['BULLET_CAMERA']
        dvr = request.form['DVR']
        remarks = request.form['REMARKS']
        data_date = request.form['DATA_DATE']

        # Check if SHOP_ID already exists
        cursor.execute('SELECT COUNT(*) FROM HARDWARE_DATA WHERE SHOP_ID = ?', (shopid,))
        count = cursor.fetchone()[0]

        if count > 0:
            flash(f'INSERT ERROR: SHOP_ID {shopid} ALREADY EXISTS', 'error')
        else:
            # Insert a new record into the HARDWARE_DATA table
            cursor.execute('''
                INSERT INTO HARDWARE_DATA (SHOP_ID, SHOP_NAME, SYSTEM, PRINTER, SCANNER,DOME_CAMERA,BULLET_CAMERA,DVR,REMARKS, DATA_DATE)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (shopid, shopname, system, printer, scanner,dome_camera,bullet_camera,dvr,remarks, data_date))
            flash(f'INSERTED SUCCESSFULLY.', 'error')
        conn.commit()
        return redirect(url_for('index'))

@app.route('/export', methods=['POST'])
def export_data():
    if request.method == 'POST':
        start_date = request.form['FROM_DATE']
        end_date = request.form['TO_DATE']

        # SQL query to retrieve data within the specified date range
        query = f"""
            SELECT *
            FROM HARDWARE_DATA
            WHERE DATA_DATE >= ? AND DATA_DATE <= ?;
        """

        cursor.execute(query, (start_date, end_date))
        results = cursor.fetchall()

        # Export data to a CSV file
        with open('exported_data.csv', 'w', newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([column[0] for column in cursor.description])
            csv_writer.writerows(results)

        return send_file('exported_data.csv', as_attachment=True)
    
@app.route('/delete', methods=['POST'])
def delete_data():
    if request.method == 'POST':
        shopid = request.form['SHOP_ID']

        # Delete the record with the specified SHOPID
        cursor.execute('''
            DELETE FROM HARDWARE_DATA
            WHERE SHOP_ID = ?
        ''', (shopid,))
        
        # Check the number of rows affected
        rows_affected = cursor.rowcount

        if rows_affected > 0:
            flash('DELETED SUCCESSFULLY!', 'success')
        else:
            flash('DELETE ERROR: SHOPID NOT FOUND', 'error')

        conn.commit()
        return redirect(url_for('index'))
    
    @app.route('/')
    def index():
     return render_template('index.html')

# Route to handle form submission
@app.route('/search', methods=['POST'])
def search():
    shop_id = request.form.get('shop_id')

    # Use the existing cursor, no need to create another one
    cursor.execute("SELECT * FROM hardware_data WHERE shop_id=?", (shop_id,))
    result = cursor.fetchone()

    return render_template('index.html', result=result)

if __name__ == '__main__':
    app.run(port=5002, debug=True)
