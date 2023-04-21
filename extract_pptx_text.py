import os
import sys
from pptx import Presentation

def extract_text_from_slide(slide):
    text = []
    for shape in slide.shapes:
        if hasattr(shape, 'text'):
            text.append(shape.text)
    return '\n'.join(text)

def extract_notes_from_slide(slide):
    notes_slide = slide.notes_slide
    if notes_slide is not None:
        return extract_text_from_slide(notes_slide)
    return ""

def main(pptx_file, output_file):
    if not os.path.exists(pptx_file):
        print(f"Error: File '{pptx_file}' does not exist.")
        sys.exit(1)

    presentation = Presentation(pptx_file)

    with open(output_file, 'w', encoding='utf-8') as f:
        for slide_num, slide in enumerate(presentation.slides, start=1):
            slide_text = extract_text_from_slide(slide)
            slide_notes = extract_notes_from_slide(slide)
            f.write(f"Slide {slide_num}:\n{slide_text}\n\nNotes:\n{slide_notes}\n\n")
            print(f"Slide {slide_num} and its notes extracted.")

    print(f"Extraction completed. Results saved in '{output_file}'.")

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python extract_pptx_text.py <pptx_file> <output_file>")
        sys.exit(1)

    pptx_file = sys.argv[1]
    output_file = sys.argv[2]
    main(pptx_file, output_file)
