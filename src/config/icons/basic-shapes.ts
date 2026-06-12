import { type CategoryConfig } from '../../types/iconConfig';

export const basicShapes: CategoryConfig = {
  category: "Basisformen (TS)",
  icons: [
    {
      name: "Kreis",
      keywords: ["rund", "ellipse", "punkt", "circle"],
      svg: `
        <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
          <circle cx="50" cy="50" r="50" fill="currentColor"/>
        </svg>
      `
    },

    {
      name: "Dreieck",
      keywords: ["triangle", "spitz", "geometrie"],
      svg: `
        <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
          <polygon points="50,10 90,90 10,90" fill="currentColor"/>
        </svg>
      `
    }
  ]
};