export interface MainIconConfig {
  version: string;
  description?: string;
  files: string[];
}

export interface IconElement {
  name: string;
  keywords: string[];
  svg: string | string[];
}

export interface CategoryConfig {
  category: string;
  icons: IconElement[];
}