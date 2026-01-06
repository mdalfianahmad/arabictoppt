
import pptxgen from 'pptxgenjs';
import { ExcelRow, ColumnMapping, IndexPosition, AspectRatio } from '../types';

export const generatePptx = async (
  data: ExcelRow[], 
  mapping: ColumnMapping, 
  fileName: string,
  allColumns: string[],
  includeIndex: boolean,
  indexPosition: IndexPosition,
  indexCol1: string,
  indexCol2: string,
  indexFormat: string,
  aspectRatio: AspectRatio,
  slideWidth: number,
  slideHeight: number
): Promise<void> => {
  const pptx = new pptxgen();
  
  pptx.defineLayout({
    name: 'CUSTOM_LAYOUT',
    width: slideWidth / 96,
    height: slideHeight / 96
  });
  pptx.layout = 'CUSTOM_LAYOUT';

  data.forEach((row) => {
    const slide = pptx.addSlide();

    Object.entries(mapping).forEach(([key, style]) => {
      if (!style.column || !row[style.column]) return;

      const text = String(row[style.column]);
      const isBox1 = key === 'box1';
      const isBox2 = key === 'box2';
      const isBox3 = key === 'box3';

      slide.addText(text, {
        x: `${style.x}%`,
        y: `${style.y}%`,
        w: `${style.w}%`,
        h: `${style.h}%`,
        fontSize: style.fontSize,
        align: 'center',
        color: '000000',
        fontFace: isBox1 ? 'Scheherazade New' : 'Arial',
        rtl: isBox1,
        bold: isBox2,
        italic: isBox3
      });
    });

    if (includeIndex) {
      const val1 = indexCol1 ? String(row[indexCol1] || '') : '';
      const val2 = indexCol2 ? String(row[indexCol2] || '') : '';
      
      let refText = indexFormat || '{1}:{2}';
      refText = refText.replace('{1}', val1).replace('{2}', val2);

      if (refText.trim()) {
        if (indexPosition === 'content-bottom-right') {
          slide.addText(refText, {
            x: '75%', y: '90%', w: '20%', h: '5%',
            fontSize: 12, fontFace: 'Arial', align: 'right', color: '888888'
          });
        } else {
          slide.addText(refText, {
            x: '2%', y: '94%', w: '96%', h: '5%',
            fontSize: 10, fontFace: 'Arial', align: 'right', color: 'AAAAAA'
          });
        }
      }
    }
  });

  const outputName = fileName.replace(/\.[^/.]+$/, "") + "_slides.pptx";
  await pptx.writeFile({ fileName: outputName });
};
