declare module "pdfmake/build/pdfmake" {
  const pdfMake: {
    addVirtualFileSystem(vfs: Record<string, string>): void;
    createPdf(docDefinition: Record<string, unknown>): {
      download(filename?: string): void;
      getBuffer(callback: (buffer: Uint8Array) => void): void;
    };
  };
  export default pdfMake;
}

declare module "pdfmake/build/vfs_fonts" {
  const vfs: Record<string, string>;
  export default vfs;
}

declare module "html-to-pdfmake" {
  function htmlToPdfmake(
    html: string,
    options?: Record<string, unknown>,
  ): unknown;
  export default htmlToPdfmake;
}
