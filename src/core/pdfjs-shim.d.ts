/**
 * pdfjs-dist legacy build 类型声明（官方包未为 legacy 子路径提供 types）。
 */
declare module 'pdfjs-dist/legacy/build/pdf.mjs' {
  export interface TextItem {
    str: string;
    transform: number[];
    width: number;
    height: number;
    hasEOL?: boolean;
  }

  export interface TextContent {
    items: TextItem[];
  }

  export interface PDFPageProxy {
    getTextContent(): Promise<TextContent>;
    cleanup(): void;
  }

  export interface PDFDocumentProxy {
    numPages: number;
    getPage(index: number): Promise<PDFPageProxy>;
    destroy(): Promise<void>;
  }

  export interface PDFDocumentLoadingTask {
    promise: Promise<PDFDocumentProxy>;
  }

  export interface GetDocumentParams {
    data: Uint8Array;
    useWorkerFetch?: boolean;
    isEvalSupported?: boolean;
    disableFontFace?: boolean;
  }

  export function getDocument(params: GetDocumentParams): PDFDocumentLoadingTask;
  export const GlobalWorkerOptions: { workerSrc: string };
  export const version: string;
}
