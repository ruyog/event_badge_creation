# event_badge_creation

Prerequisite is Windows environment with MS-Office applications installed, especially, MS-Powerpoint.

## Intermediate Badge (Badge in PPT)
 * Replace the input file in edit_ppt_badge.py at #109 with actual absolute location

 * Replace the id template (type of event) absolute location at line 102 (pptx_file)
 * Replace the intermediate file location in edit_ppt_badge.py that reflect the appropriate event location, i.e., separate folder for event_1, event_2, event_3, etc.,
 
 * Replace the header in Input_File_With_Data with actuals.
 * Edit #13 in edit_ppt_badge.py to reflect the tab name 'Event_1'/'Event_2'

 * Execute the edit_ppt_badge.py script

### Final Badge (Badge in PDF)
 * Replace the input/output directory for approriate events in ppt_to_pdf.py at #38 and #39
 * Execute the script ppt_to_pdf.py
