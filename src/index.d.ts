// index.d.ts

declare module "html-to-docx" {
    /**
     * Options pour la conversion HTML en DOCX.
     */
    interface HtmlToDocxOptions {
      /**
       * Définit le nom du fichier DOCX généré.
       */
      fileName?: string;
  
      /**
       * Définit les marges du document.
       * Exemple : { top: 1, bottom: 1, left: 1, right: 1 }
       */
      margins?: {
        top?: string;
        bottom?: string;
        left?: string;
        right?: string;
      };
    }
  
    /**
     * Convertit un contenu HTML en un fichier DOCX.
     *
     * @param html - Le contenu HTML à convertir.
     * @param options - Les options de configuration pour la conversion.
     * @returns Une promesse qui se résout en un `ArrayBuffer` contenant les données du fichier DOCX.
     */
    function htmlToDocx(
      html: string,
      options?: HtmlToDocxOptions
    ): Promise<ArrayBuffer>;
  
    export default htmlToDocx;
  }
  