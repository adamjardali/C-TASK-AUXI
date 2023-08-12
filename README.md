# TASK-AUXI
This is the solution of the task
## Usage

### PowerPointAnalyzer

The `PowerPointAnalyzer` component allows you to analyze and extract information from PowerPoint presentations. It provides the following features:

#### Get Slide Properties

You can use the `get_slide_properties` method to retrieve various properties of a slide, such as its title, dimensions, and background color. Here's an example of how to use it:

```python
from powerpoint_analyzer import PowerPointAnalyzer

# Initialize PowerPointAnalyzer with the presentation file name
presentation_file = "presentation.pptx"
analyzer = PowerPointAnalyzer(presentation_file)

# Get slide properties by slide number
slide_number = 1
slide_properties = analyzer.get_slide_properties(slide_number)

# Print slide properties
print(slide_properties)

