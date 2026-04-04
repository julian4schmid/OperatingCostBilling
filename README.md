# Automated Operating Cost Billing (work in progress)

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
│   ├── on_demand_import.py     # Import building and tenant data 
│   ├── yearly_import.py        # Import yearly cost data
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