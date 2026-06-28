# Project Intelligence Hub

## Time Impact Analysis Report | Director Pack

The Director Pack generator uses the master Word template in:

- `templates/time_impact_analysis_report_director_pack.docx`

Generation workflow:

1. Open the existing Streamlit app.
2. Go to `Delay TIA -> Download Reports`.
3. Open the `Report Generator` section.
4. Complete the required fields:
   - Project Name
   - Data Date
   - Revision
5. Review the replacement preview and source status.
6. Click `Generate Time Impact Analysis Report | Director Pack`.
7. Download the generated DOCX.

Current behavior:

- The original DOCX template is copied before any replacement.
- Placeholders and selected table values are updated only.
- Original charts and images are preserved by default.
- Every generated DOCX is logged in the SQLite `generated_outputs` table.

Current limitation:

- Embedded charts and images are preserved by default and are not regenerated unless future chart-replacement support is added.

Supported replacement fields:

- Project name
- Contract number
- Data date
- Revision
- Employer / Client
- Contractor
- Contract form / clause basis
- Accepted baseline programme
- Impacted update programme
- Calendar basis
- Schedule file name
- Schedule options
- Longest path / critical path basis
- Retained logic / progress override
- Out-of-sequence progress treatment
- Constraints
- Calendars
- Open ends
- Negative float
