# Data Migration Scripts (Python)
This repository contains Python scripts I developed during a complex data migration project involving job cost transactions and ERP integration.

The project required flexibility and adaptability, as source files were frequently updated and originated from various enterprise units, each using different formats and column structures in Excel.

Key responsibilities included:

- Consolidating and standardizing multiple mapping files with inconsistent schemas.
- Building logic to dynamically handle varying job cost transaction scenarios.
- Developing reusable, robust scripts to process and map transactions based on their context and mapping rules.
- Delivering fully mapped and validated Excel outputs ready for upload into the new ERP system.
- This project showcases my ability to:
- Work with messy real-world enterprise data.
- Adapt quickly to evolving requirements.
- Build scalable and reliable data pipelines under time-sensitive conditions.

Scripts:
- `job_mapper.py`: Handles mapping for two types of jobs: active jobs and balance forwarding jobs.
For active jobs (Type 1), all transactions are mapped row by row based on specific formatting rules for the output file.
For balance forwarding jobs (Type 3), the script generates a summarized output according to a different set of defined rules.
The script supports both job types and applies the appropriate mapping logic for each.