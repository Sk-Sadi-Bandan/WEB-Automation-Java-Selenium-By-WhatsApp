1. Install the Tesseract software from the exe folder.
2. Add the Tesseract path in the Environment variable. (If faceing any error or not understand properly please watch [this ](https://www.youtube.com/watch?v=LMM6s9JL5GA&t=1s)video )
3. Create a python environment.

   `python -m venv env`
4. Active the environment

   `.\env\Scripts\Activate.ps1`

   if show any error , Start powershall run as administrator

   enter this comman
   `set-executionpolicy remotesigned`

   type Y for confirmation
5. Then go to VScode and Active the environment
6. Install all libraries

   `pip install -r requirements.txt`
7. Put all the images in piture folder
8. Run the code

   `python main.py`
9. Num will save in phone_numbers.txt file
