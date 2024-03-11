import time
import comtypes.client
import os


def pptx_to_pdf_directory(input_dir, output_dir):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = True  # Set to True to make PowerPoint visible

    # Iterate through all files in the input directory
    for filename in os.listdir(input_dir):
        if filename.endswith(".pptx"):
            pptx_file = os.path.join(input_dir, filename)
            output_pdf = os.path.join(output_dir, filename.replace(".pptx", ".pdf"))

            # Open the PowerPoint file and save it as PDF
            presentation = powerpoint.Presentations.Open(pptx_file)
            presentation.SaveAs(output_pdf, 32)  # 32 represents the PDF format
            try:
                presentation.Close()
            except Exception as e:
                print(filename, e)


    powerpoint.Quit()
    time.sleep(2)


#input_dir = 'C:\\Users\\msridhar20\\PycharmProjects\\scopus_data_processing\\badges\\intermediate_conference_badges'
#output_dir = 'C:\\Users\\msridhar20\\PycharmProjects\\scopus_data_processing\\badges\\final_conference_badges'

#input_dir = 'C:\\Users\\msridhar20\\PycharmProjects\\scopus_data_processing\\badges\\intermediate_conf_workshop_badges'
#output_dir = 'C:\\Users\\msridhar20\\PycharmProjects\\scopus_data_processing\\badges\\final_conf_workshop_badges'

#input_dir = 'C:\\Users\\msridhar20\\PycharmProjects\\scopus_data_processing\\badges\\intermediate_workshop_badges'
#output_dir = 'C:\\Users\\msridhar20\\PycharmProjects\\scopus_data_processing\\badges\\final_workshop_badges'

input_dir = 'C:\\Users\\msridhar20\\PycharmProjects\\scopus_data_processing\\badges\\intermediate_organizer_badges'
output_dir = 'C:\\Users\\msridhar20\\PycharmProjects\\scopus_data_processing\\badges\\final_organizer_badges'

#input_dir = 'C:\\Users\\msridhar20\\PycharmProjects\\scopus_data_processing\\badges\\intermediate_guest_badges'
#output_dir = 'C:\\Users\\msridhar20\\PycharmProjects\\scopus_data_processing\\badges\\final_guest_badges'
pptx_to_pdf_directory(input_dir, output_dir)
