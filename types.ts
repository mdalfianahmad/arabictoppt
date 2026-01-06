
export interface ExcelRow {
  [key: string]: any;
}

export interface ColumnStyle {
  column: string;
  fontSize: number;
  x: number; // percentage
  y: number; // percentage
  w: number; // percentage
  h: number; // percentage
}

export interface ColumnMapping {
  box1: ColumnStyle;
  box2: ColumnStyle;
  box3: ColumnStyle;
  box4: ColumnStyle;
  box5: ColumnStyle;
}

export enum AppStep {
  INTRO = 'INTRO',
  UPLOAD = 'UPLOAD',
  MAPPING = 'MAPPING',
  PREVIEW = 'PREVIEW',
  DOWNLOAD = 'DOWNLOAD'
}

export type RowSelectionMode = 'all' | 'range' | 'specific';
export type IndexPosition = 'content-bottom-right' | 'footer';
export type AspectRatio = '16:9' | '4:3' | 'Custom';

export interface AppState {
  step: AppStep;
  file: File | null;
  data: ExcelRow[];
  columns: string[];
  mapping: ColumnMapping;
  isGenerating: boolean;
  error: string | null;
  selectionMode: RowSelectionMode;
  rangeStart: number;
  rangeEnd: number;
  specificRows: string;
  includeIndex: boolean;
  indexPosition: IndexPosition;
  indexCol1: string;
  indexCol2: string;
  indexFormat: string;
  aspectRatio: AspectRatio;
  slideWidth: number;
  slideHeight: number;
}
