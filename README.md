# Automated Operating Cost Billing

## Overview
OperatingCostBilling is a Python-based project for managing and calculating operating costs for rental properties.

It uses a PostgreSQL database to store structured data such as buildings, units, tenants, and cost information. The goal is to automate the process of creating operating cost statements (Nebenkostenabrechnungen), including data import, calculation, and report generation.

The project is designed with a clean separation between data storage (SQL), processing (Python), and output (Excel/PDF).

---


## Project Structure

```text
OperatingCostBilling/
├── scripts/                    
│   ├── backup_db.py            # Backup PostgreSQL database
│   ├── import_excel.py         # Import cost and tenant data from Excel
│   ├── calculate_billing.py    # Calculate costs per tenant/unit
│   ├── generate_reports.py     # Generate Excel/PDF reports
│   └── config.py               # Local database credentials
├── data/                       
│   ├── templates/              
│   └── backups/                
├── .gitignore                  
└── README.md         
```


---

## Author
Julian Schmid

https://github.com/julian4schmid