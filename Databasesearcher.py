import pyodbc
from tkinter import Tk, filedialog
from threading import Thread, Lock
import signal
import sys
import time

# Global variable and lock for progress tracking
progress = {'current': 0, 'total': 0}
progress_lock = Lock()
process_thread = None
running = True

def signal_handler(sig, frame):
    global running
    print("\nInterrupt received. Exiting...")
    running = False
    if process_thread and process_thread.is_alive():
        # Signal to stop the processing
        process_thread.join()  # Wait for the thread to finish if necessary
    sys.exit(0)

def connect_to_database(database_file):
    try:
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            rf'DBQ={database_file};'
        )
        conn = pyodbc.connect(conn_str)
        return conn
    except pyodbc.Error as e:
        print(f"Error connecting to the database: {e}")
        return None

def get_table_names(conn):
    try:
        cursor = conn.cursor()
        table_names = [row.table_name for row in cursor.tables(tableType='TABLE')]
        return table_names
    except pyodbc.Error as e:
        print(f"Error retrieving table names: {e}")
        return None

def query_table(conn, table_name, chunk_size=1000):
    query = f"SELECT * FROM [{table_name}]"
    try:
        cursor = conn.cursor()
        cursor.execute(query)
        while True:
            if not running:
                break
            rows = cursor.fetchmany(chunk_size)
            if not rows:
                break
            yield rows
    except pyodbc.Error as e:
        print(f"Error querying table: {e}")

def process_results(database_file, table_name, results_holder):
    global progress
    conn = connect_to_database(database_file)
    if conn:
        # Estimate total rows for progress calculation
        cursor = conn.cursor()
        cursor.execute(f"SELECT COUNT(*) FROM [{table_name}]")
        total_rows = cursor.fetchone()[0]
        progress['total'] = total_rows

        results = []
        for chunk in query_table(conn, table_name):
            with progress_lock:
                progress['current'] += len(chunk)
            results.extend(chunk)
        conn.close()
        results_holder.append(results)
    else:
        results_holder.append([])

def save_results_to_txt(results, filename):
    with open(filename, 'w') as file:
        for row in results:
            file.write(str(row) + '\n')

def show_progress(event=None):
    with progress_lock:
        if progress['total'] > 0:
            percentage = (progress['current'] / progress['total']) * 100
            print(f"Scanning Progress: {percentage:.2f}%")
        else:
            print("Scanning Progress: Not started or no data available.")

def main():
    global progress, process_thread, running
    root = Tk()
    root.withdraw()  # Hide the main window

    # Set up signal handling
    signal.signal(signal.SIGINT, signal_handler)

    # Set up keyboard shortcuts for Ctrl + S
    root.bind_all('<Control-s>', show_progress)

    # Ask user to select database file
    database_file = filedialog.askopenfilename(title="Select Access database file", filetypes=[("Access Database", "*.mdb;*.accdb")])
    if not database_file:
        print("No database file selected. Exiting.")
        return

    conn = connect_to_database(database_file)
    if conn:
        # Get and display table names
        table_names = get_table_names(conn)
        if table_names:
            print("Table names:")
            for idx, name in enumerate(table_names):
                print(f"{idx + 1}. {name}")

            # Ask user to select a table
            table_index = input("Enter the number of the table you want to query: ")
            try:
                table_index = int(table_index) - 1
                if 0 <= table_index < len(table_names):
                    table_name = table_names[table_index]

                    # Create a list to hold the search results
                    results_holder = []

                    # Start a thread to process the table
                    process_thread = Thread(target=process_results, args=(database_file, table_name, results_holder))
                    process_thread.start()

                    # Monitor the thread and allow for `Ctrl + C` handling
                    while process_thread.is_alive() and running:
                        time.sleep(1)

                    # Wait for the thread to complete
                    process_thread.join()

                    # Get the search results
                    results = results_holder[0]

                    if results:
                        print("Search results:")
                        for row in results:
                            print(row)

                        # Ask user if they want to save the results to a file
                        save_to_file = input("Do you want to save the results to a file? (yes/no): ").lower()
                        if save_to_file == 'yes':
                            output_file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
                            if output_file:
                                save_results_to_txt(results, output_file)
                                print(f"Results saved to {output_file}")
                else:
                    print("Invalid table number selected.")
            except ValueError:
                print("Invalid input. Please enter a number.")
        else:
            print("No tables found in the database.")
        conn.close()
    else:
        print("Failed to connect to the database.")

    # Start the Tkinter event loop to capture keyboard events
    root.mainloop()

if __name__ == "__main__":
    main()
