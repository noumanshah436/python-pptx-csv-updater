**New Tech Challenge For Interview**

**Requirements:**
- Github account.
- Google (drive) account.

**Test:**
1. Create a Google presentation with 2 slides.
   - The first slide should contain an image with 2 text fields next to it (Title and description).
   - The title field should be bigger than the description field.
   - Their font colors should also be different.
   - The second slide should have 2 images and one text field.
2. Add keys to the 6 objects created above and insert the key as “alt text” for the 6 objects.
3. Download the Google presentation as a pptx file.
4. Fork the [python-pptx repository](https://github.com/scanny/python-pptx).
5. Apply the following PR to your new repository: [PR #512](https://github.com/scanny/python-pptx/pull/512/files).
6. Create a CSV file with the following columns:
   - Key
   - Data
7. For each key in the alt text in the pptx, match a data value in the CSV. For text fields, the data should be text, and for images, it should be a local image. Ensure all text fields and images are unique.
8. Create a Python script that:
   - Imports the python-pptx library.
   - Loads the CSV and then loads the pptx file.
   - Verifies that the pptx has all CSV keys.
   - Verifies that the CSV has the correct data type for each field.
   - Replaces the images and text in the pptx file based on the alt text keys and the new data in the CSV.
   - Saves the updated pptx file.
