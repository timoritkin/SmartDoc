import sqlite3
import time


# Validate the ID
def validate_input(patient_id, phone):
    # Check if ID is a 9-digit integer
    # if len(patient_id) != 9 or not patient_id.isdigit():
    #     raise ValueError("תעודת זהות אמורה להכיל 9 מספרים")

    if len(phone) != 10 or not phone.isdigit():
        raise ValueError("מספר טלפון אמור להכיל 10 מספרים")

    return True


def create_tables(db_path):
    # Connect to SQLite database (or create it if it doesn't exist)
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Enable foreign key constraint enforcement (important for SQLite)
    cursor.execute("PRAGMA foreign_keys = ON")

    # Check if the 'patients' table already exists
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='patients'")
    if cursor.fetchone() is None:
        # If the table doesn't exist, create it
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS patients (
            patient_id INTEGER (9) PRIMARY KEY,
            first_name VARCHAR  (50) NOT NULL,
            last_name VARCHAR (50) NOT NULL,
            birthdate VARCHAR (11) NOT NULL,
            phone_number TEXT NOT NULL CHECK (LENGTH(phone_number) = 10)
        )
        """)

    # Check if the 'visits' table already exists
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='visits'")
    if cursor.fetchone() is None:
        # If the table doesn't exist, create it
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS visits (
            visit_id INTEGER PRIMARY KEY AUTOINCREMENT,
            patient_id INTEGER NOT NULL,
            visit_date VARCHAR (11) NOT NULL,
            docx_path TEXT NOT NULL,
            FOREIGN KEY (patient_id) REFERENCES patients(patient_id)
            ON DELETE CASCADE ON UPDATE CASCADE -- Maintain referential integrity
            
        )
        """)

    # Commit the changes and close the connection
    conn.commit()
    conn.close()


def fetch_visit_data(db_path):
    """Fetch data from the SQLite database."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Query to fetch patient and visit details (without docx_path)
    query = """
        SELECT v.visit_date,p.phone_number, p.birthdate,  p.first_name,p.last_name, p.patient_id
        FROM patients p
        LEFT JOIN visits v ON p.patient_id = v.patient_id
    """
    cursor.execute(query)
    rows = cursor.fetchall()

    conn.commit()
    conn.close()
    return rows


def fetch_patient_data(db_path):
    """Fetch data from the SQLite database."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Query to fetch patient and visit details (without docx_path)
    query = """
        SELECT p.phone_number, p.birthdate,  p.first_name,p.last_name, p.patient_id
        FROM patients p
       
    """
    cursor.execute(query)
    rows = cursor.fetchall()

    conn.commit()
    conn.close()
    return rows


def insert_visit_record(ID, current_date, docx_path, db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    try:
        # Your insert logic here
        cursor.execute("INSERT INTO visits (patient_id, visit_date, docx_path) VALUES (?, ?, ?)",
                       (ID, current_date, docx_path))

        # Commit the transaction
        conn.commit()
    except Exception as e:
        print(f"Error: {e}")
    finally:
        # Ensure the connection is closed
        conn.close()


def insert_patient_record(first_name, last_name, patient_id, birthdate, phone, db_path):
    """ Inserts a new patient record into the SQLite database. """
    # Validate the patient ID
    validate_input(patient_id, phone)  # This will raise an error if the ID is invalid

    # Connect to the SQLite database
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Insert the patient record into the patients table
    cursor.execute("""
    INSERT INTO patients (patient_id, first_name, last_name, birthdate, phone_number)
    VALUES (?, ?, ?, ?, ?)
    """, (patient_id, first_name, last_name, birthdate, phone))

    # Commit the changes and close the connection
    conn.commit()
    conn.close()


def get_docx_path(patient_id, visit_date, db_path):
    """
       Retrieve the docx path for a specific patient, optionally filtered by visit date or visit ID.
       """
    import sqlite3

    # Connect to SQLite database
    conn = sqlite3.connect(db_path)
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

    conn.commit()
    # Close the connection
    conn.close()

    if result:
        return result[0]  # Return the file path
    else:
        return None  # Return None if no matching record is found


def search_patients_visits(search_term, db_path):
    """
    Search patients in the database based on a search term.

    :param search_term: String to search for in patient records
    :return: List of matching patient records
    """
    # Connect to the database
    conn = sqlite3.connect(db_path)
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

    conn.commit()
    conn.close()

    return results


def search_patients_data(search_term, db_path):
    """
    Search patients in the database based on a search term.

    :param search_term: String to search for in patient records
    :return: List of matching patient records
    """
    # Connect to the database
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Create a search query that checks multiple columns
    query = """
        SELECT p.phone_number, p.birthdate,  p.first_name,p.last_name, p.patient_id
        FROM patients p

        WHERE 
            LOWER(p.first_name) LIKE ? OR 
            LOWER(p.last_name) LIKE ? OR 
            LOWER(p.birthdate) LIKE ? OR 
            LOWER(p.patient_id) LIKE ?
        """

    # Use % wildcards for partial matching
    search_param = f'%{search_term.lower()}%'

    # Execute the query
    cursor.execute(query, (search_param, search_param, search_param, search_param))

    # Fetch and process results
    results = cursor.fetchall()

    conn.commit()
    conn.close()

    return results


# Function to check if patient_id exists in the database
def check_patient_id_exists(patient_id, db_path):
    # Connect to the SQLite database (change 'your_database.db' to your database file)
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Prepare the query to check if patient_id exists
    query = "SELECT * FROM patients WHERE patient_id = ?"

    # Execute the query with the patient_id as parameter
    cursor.execute(query, (patient_id,))

    # Fetch one result; if no result is found, None will be returned
    result = cursor.fetchone()

    conn.commit()
    # Close the connection
    conn.close()

    # Return True if the patient_id exists, False otherwise
    return result is not None


def update_patient_record(fields, old_patient_id, tree, db_path, max_retries=5):
    conn = None
    for attempt in range(max_retries):
        try:
            # Close any existing connection before attempting a new one
            if conn:
                try:
                    conn.close()
                except:
                    pass

            # Create a new connection with a longer timeout
            conn = sqlite3.connect(db_path, timeout=30)
            cursor = conn.cursor()

            # Enable immediate transaction mode but with error handling
            try:
                conn.execute("BEGIN IMMEDIATE")
            except sqlite3.OperationalError as e:
                if "database is locked" in str(e):
                    # If locked, close connection, wait longer, and retry
                    conn.close()
                    time.sleep(2 * (attempt + 1))  # Exponential backoff
                    continue
                else:
                    raise

            # Extract updated values
            new_data = {}
            for key in fields:
                if key == "גיל":
                    try:
                        if hasattr(fields[key], 'get_date'):
                            new_data[key] = fields[key].get_date().strftime('%d/%m/%Y')
                        else:
                            new_data[key] = fields[key].get()
                    except Exception:
                        new_data[key] = ""
                else:
                    new_data[key] = fields[key].get()

            # Extract individual values
            first_name = new_data["שם פרטי"]
            last_name = new_data["שם משפחה"]
            new_patient_id = new_data["תעודת זהות"]
            phone = str(new_data["טלפון"])
            birth_date = new_data.get("גיל", "")

            # If patient ID has changed, update the visits table first
            if str(new_patient_id) != str(old_patient_id):
                cursor.execute("""
                    UPDATE visits 
                    SET patient_id = ? 
                    WHERE patient_id = ?
                """, (new_patient_id, old_patient_id))

            # Then update the patient record
            cursor.execute("""
                UPDATE patients 
                SET patient_id = ?, first_name = ?, last_name = ?, birthdate = ?, phone_number = ?
                WHERE patient_id = ?
            """, (new_patient_id, first_name, last_name, birth_date, phone, old_patient_id))

            conn.commit()
            conn.close()
            return True  # Success

        except sqlite3.OperationalError as e:
            if "database is locked" in str(e) and attempt < max_retries - 1:
                # Wait longer between retries with exponential backoff
                wait_time = 2 * (attempt + 1)
                time.sleep(wait_time)
                continue
            else:
                # Close the connection even if there's an error
                if conn:
                    try:
                        conn.close()
                    except:
                        pass
                print(f"Database error after {attempt + 1} attempts: {str(e)}")
                raise  # Re-raise the exception
        except Exception as e:
            # Handle any other exceptions
            if conn:
                try:
                    conn.close()
                except:
                    pass
            print(f"Unexpected error: {str(e)}")
            raise

    # If we've exhausted all retries
    raise sqlite3.OperationalError("Failed to update patient record after maximum retries")
def get_patient_birthdate(patient_id, db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Fetch birthdate string from DB
    cursor.execute("SELECT birthdate FROM patients WHERE patient_id = ?", (patient_id,))
    result = cursor.fetchone()

    conn.close()
    return result


def get_patient_docx_path(patient_id, db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    print(patient_id)
    # Fetch birthdate string from DB
    cursor.execute("SELECT docx_path FROM visits WHERE patient_id = ?", (patient_id,))
    result = cursor.fetchone()

    conn.close()
    return result
