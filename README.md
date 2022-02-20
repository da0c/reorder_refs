# reorder_refs

This is an utility for reordering references in .docx manuscripts or reports.  
If you don't use something like EndNote this is for you.  
  
**Usage:**  
```
python reorder_refs.py input.docx output.docx rules.conf  
```  
After this you get output.docx with reordered references.
Each processed references will be marked with *, like [*1]
**Important** - see the console output for reordered references list.

See the `rules.conf` sample for configuration details.  

**Sample input text**  

Sample text for testing the reoprder_refs.pyutility [1]
Many teams have been actively pursuing research into the lightweight high-resolution optics for remote sensing needs [MP1-MP3]. One of the first working prototypes using 20-meter diameter ultralight diffractive film lens as a high-resolution satellite imaging system was a DARPA project called “Membrane Optical Imager for Real-Time Exploitation” initiated in 2012 [5], [1] but resulted in passing the on-the-ground tests only in 2014 [MP2]. Despite the lack of production implementation of a telescope based on ultra-light diffraction elements, these studies [7] are still relevant and contribute to the progress in this area.  


**Output text**  

Sample text for testing the reoprder_refs.py utility [*1]
Many teams have been actively pursuing research into the lightweight high-resolution optics for remote sensing needs [*2],[*3],[*4]. One of the first working prototypes using 20-meter diameter ultralight diffractive film lens as a high-resolution satellite imaging system was a DARPA project called “Membrane Optical Imager for Real-Time Exploitation” initiated in 2012 [*5], [*1] but resulted in passing the on-the-ground tests only in 2014 [*3]. Despite the lack of production implementation of a telescope based on ultra-light diffraction elements, these studies [*6] are still relevant and contribute to the progress in this area.  
  
**Reordered refs (from console output):**

{'1': 1, 'MP1': 2, 'MP2': 3, 'MP3': 4, '5': 5, '7': 6}

