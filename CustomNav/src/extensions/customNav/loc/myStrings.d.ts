/**
 * Type declaration for the generated localization string bundle used by the
 * CustomNav application customizer.
 */
declare interface ICustomNavApplicationCustomizerStrings {
  /** Title used by the extension where string-based titles are required. */
  Title: string;
}

/** Module declaration that lets TypeScript import the generated string bundle. */
declare module 'CustomNavApplicationCustomizerStrings' {
  const strings: ICustomNavApplicationCustomizerStrings;
  export = strings;
}
