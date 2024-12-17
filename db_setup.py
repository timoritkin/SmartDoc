import sqlite3


def create_tables():
    # Connect to SQLite database (or create it if it doesn't exist)
    conn = sqlite3.connect("patients.db")
    cursor = conn.cursor()

    # Enable foreign key constraint enforcement (important for SQLite)
    cursor.execute("PRAGMA foreign_keys = ON")

    # Check if the 'patients' table already exists
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='patients'")
    if cursor.fetchone() is None:
        # If the table doesn't exist, create it
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS patients (
            patient_id TEXT PRIMARY KEY,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            birthdate TEXT NOT NULL,
            phone_number TEXT 
        )
        """)


    # Check if the 'visits' table already exists
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='visits'")
    if cursor.fetchone() is None:
        # If the table doesn't exist, create it
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS visits (
            visit_id INTEGER PRIMARY KEY AUTOINCREMENT,
            patient_id TEXT NOT NULL,
            visit_date TEXT NOT NULL,
            docx_path TEXT NOT NULL,
            FOREIGN KEY (patient_id) REFERENCES patients(patient_id)
            
        )
        """)

    # Commit the changes and close the connection
    conn.commit()
    conn.close()


def fetch_data():
    """Fetch data from the SQLite database."""
    conn = sqlite3.connect("patients.db")
    cursor = conn.cursor()

    # Query to fetch patient and visit details (without docx_path)
    query = """
        SELECT v.visit_date,p.phone_number, p.birthdate,  p.first_name,p.last_name, p.patient_id
        FROM patients p
        LEFT JOIN visits v ON p.patient_id = v.patient_id
    """
    cursor.execute(query)
    rows = cursor.fetchall()


    conn.close()
    return rows


def insert_visit_record(patient_id, time, docx_path):
    # Connect to the SQLite database
    conn = sqlite3.connect("patients.db")
    cursor = conn.cursor()

    # Insert the patient record into the patients table
    cursor.execute("""
        INSERT INTO visits (patient_id, visit_date, docx_path)
        VALUES (?, ?, ?)
        """, (patient_id, time, docx_path,))

    # Commit the changes and close the connection
    conn.commit()
    conn.close()


def insert_patient_record(first_name, last_name, patient_id, birthdate, phone):
    """
    Inserts a new patient record into the SQLite database.

    Parameters:
        first_name (str): The patient's first name.
        last_name (str): The patient's last name.
        patient_id (str): The patient's unique ID.
        birthdate (str): The patient's age (as a string).
        time (str): The time of the record (e.g., a timestamp or date).
        docx_path (str): The path to the patient's document file.
        :param phone:
    """
    # Connect to the SQLite database
    conn = sqlite3.connect("patients.db")
    cursor = conn.cursor()

    # Insert the patient record into the patients table
    cursor.execute("""
    INSERT INTO patients (patient_id, first_name, last_name, birthdate,phone_number)
    VALUES (?, ?, ?, ?,?)
    """, (patient_id, first_name, last_name, birthdate, phone))

    # Commit the changes and close the connection
    conn.commit()
    conn.close()


def get_docx_path(patient_id, visit_date):
    """
       Retrieve the docx path for a specific patient, optionally filtered by visit date or visit ID.
       """
    import sqlite3

    # Connect to SQLite database
    conn = sqlite3.connect("patients.db")
    cursor = conn.cursor()

    # Base query
    query = """
       SELECT docx_path
       FROM visits
       WHERE patient_id = ?
       """
    params = [patient_id]

    # Add optional filters
    if visit_date:
        query += " AND visit_date = ?"
        params.append(visit_date)



    # Execute the query
    cursor.execute(query, params)

    # Fetch the result
    result = cursor.fetchone()

    # Close the connection
    conn.close()

    if result:
        return result[0]  # Return the file path
    else:
        return None  # Return None if no matching record is found


def search_patients(search_term):
    """
    Search patients in the database based on a search term.

    :param search_term: String to search for in patient records
    :return: List of matching patient records
    """
    # Connect to the database
    conn = sqlite3.connect("patients.db")
    cursor = conn.cursor()

    # Create a search query that checks multiple columns
    query = """
        SELECT  v.visit_date,p.phone_number, p.birthdate, p.first_name, p.last_name, p.patient_id
        FROM patients p
        LEFT JOIN visits v ON p.patient_id = v.patient_id
        WHERE 
            LOWER(p.first_name) LIKE ? OR 
            LOWER(p.last_name) LIKE ? OR 
            LOWER(p.birthdate) LIKE ? OR 
            LOWER(v.visit_date) LIKE ? OR
            LOWER(p.patient_id) LIKE ?
        """

    # Use % wildcards for partial matching
    search_param = f'%{search_term.lower()}%'

    # Execute the query
    cursor.execute(query, (search_param, search_param, search_param, search_param, search_param))

    # Fetch and process results
    results = cursor.fetchall()

    conn.close()

    return results


# Function to check if patient_id exists in the database
def check_patient_id_exists(patient_id):
    # Connect to the SQLite database (change 'your_database.db' to your database file)
    conn = sqlite3.connect('patients.db')
    cursor = conn.cursor()

    # Prepare the query to check if patient_id exists
    query = "SELECT * FROM patients WHERE patient_id = ?"

    # Execute the query with the patient_id as parameter
    cursor.execute(query, (patient_id,))

    # Fetch one result; if no result is found, None will be returned
    result = cursor.fetchone()

    # Close the connection
    conn.close()

    # Return True if the patient_id exists, False otherwise
    return result is not None
