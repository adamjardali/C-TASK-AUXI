from PowerPointAnalyzer import PowerPointAnalyzer
from PowerPointManager import PowerPointManager


#Output and Input File
file_name = "testing.pptx"
output_file_name = "outputfile.pptx"

# Create PowerPointAnalyzer object and PowerPointManager Object
power_point_analyzer = PowerPointAnalyzer(file_name)
power_point_manager = PowerPointManager()

#Get the intial title of slide one input
slide_one_testing = power_point_analyzer.get_slide(1)
slide_one_title = power_point_analyzer.get_title(slide_one_testing)


#Get all the text boxes
text_boxes = power_point_analyzer.get_text_boxes_in_slide(slide_one_testing)
counter = 0

#Create Slide 1 in output and get it
power_point_manager.create_slide()
slide_one_output = power_point_manager.get_slide(1)


# Title methods => Add title, change title font, size, color, and bold.
power_point_manager.add_title(slide_one_output,slide_one_title)
power_point_manager.change_title_font(slide_one_output,"Beirut")
power_point_manager.change_title_size(slide_one_output,44)



#Get all the shapes in the input file with its properties (Color, Shape Type, width, top ...)
get_slide_one_shapes = power_point_analyzer.get_auto_shapes(slide_one_testing)


#Create all the shapes
for i in range(len(get_slide_one_shapes)):
	shape_type = get_slide_one_shapes[i]['auto_shape_type']

	text_to_add = ''

	if(counter < len(text_boxes)):
		text_to_add = text_boxes[counter]
		counter += 1

	to_insert_shape = power_point_manager.insert_auto_shape_by_type(slide_one_output,str(shape_type))
	power_point_manager.add_title_to_shape(to_insert_shape,text_to_add)


for i in range(counter,len(text_boxes)):
	text_list = text_boxes[i].split('\n')
	power_point_manager.add_text_box_with_bullet_points_to_slide(slide_one_output,text_list)


#get slides properties 
# power_point_manager.get_slide_properties(slide_one_output)
# power_point_analyzer.analyze_presentation()

#Get all the text boxes

#Save the output file.
power_point_manager.save_presentation(output_file_name)
