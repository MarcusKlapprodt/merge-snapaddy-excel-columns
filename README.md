# Merge snapADDY Excel Columns

This Python application is designed to streamline the process of modifying snapADDY VisitReport Excel downloads by merging and transforming data according to your specific needs.
It was written to help a customer. As the customer implemented a CRM system, the need to do it via excel has lost it's usefulness. Therefore, this repo is not actively maintained anymore. It could be useful and retrofitted for other use cases

## Features

- **Concatenate Columns**  
  Automatically concatenate specified columns from the snapADDY VisitReport Excel file, allowing for easier data management and analysis.

- **Reference User Sales Organization Tokens**  
  Utilize a user table to reference and insert sales organization tokens directly into the VisitReport, enhancing the report’s relevance and accuracy.

## Installation

1. **Clone the Repository**

   ```bash
   git clone https://github.com/MarcusKlapprodt/merge-snapaddy-excel-columns.git
   ```

2. **Set Up a Virtual Environment**

   ```bash
   python3 -m venv .env
   source .env/bin/activate  # On Windows use `.env\Scripts\activate`
   ```

3. **Install Required Dependencies**

   Install the necessary Python packages using pip:

   ```bash
   pip install -r requirements.txt
   ```

4. **Run the Application**

   Start the application with the following command:

   ```bash
   python snapaddy-xlsx-config.py
   ```

## Notes

- This application is specifically designed for snapADDY VisitReport files but can be adapted for other use cases with similar data structures.
- Feel free to modify and extend the code to fit your unique requirements.

## License

This project is licensed under the GNU General Public License v3.0 (GPL-3.0).

## Contributions

Contributions are welcome! Please feel free to submit a pull request or open an issue if you have suggestions for improvements or additional features.

---

This project is not actively maintained, but you are encouraged to use, adapt, and expand it as needed.
