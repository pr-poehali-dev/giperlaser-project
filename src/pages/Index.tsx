import { useState } from "react";
import Icon from "@/components/ui/icon";
import PptxGenJS from "pptxgenjs";

const IMGS = {
  laser: "https://cdn.poehali.dev/projects/edca03ca-d59d-4b05-889c-21f49114664a/files/01e369cc-f9a3-4c2e-a810-fc342bc061d3.jpg",
  machine3d: "https://cdn.poehali.dev/files/6f15724a-146b-4676-8a17-10f4d6e9f9af.png",       // 1.png — полный вид станка (рендер)
  metalParts: "https://cdn.poehali.dev/files/d1e49a08-0c46-4d82-ad1a-b3083fdcd0aa.jpg",      // 2.jpg — металлические детали (алюм.)
  blackParts: "https://cdn.poehali.dev/files/3775b244-324e-4465-9675-1d73f2f598b9.jpg",       // 3.jpg — чёрные детали
  brassClamp: "https://cdn.poehali.dev/files/f19736bf-ff2f-43d5-a16a-59ea7f18f0ca.jpg",      // 4.jpg — латунные зажимы
  portalBeam: "https://cdn.poehali.dev/files/e052853d-4003-4f78-92aa-ea331ce5a82b.jpg",      // 5.jpg — портальная балка 3D
  machineFrame: "https://cdn.poehali.dev/files/4354f62d-6d5c-4481-bdb8-c243aecfa9df.jpg",    // 6.jpg — рама станка 3D
  zAxis: "https://cdn.poehali.dev/files/58383c1a-a4b0-4b99-9e45-df240c01e837.jpg",           // 7.jpg — ось Z / суппорт 3D
  techDiagram: "https://cdn.poehali.dev/files/5fd2e328-1b4e-4d5d-8827-aa8e83e7167a.jpg",     // 8.jpg — схема режущей головы
  realMachine: "https://cdn.poehali.dev/files/51594365-a69b-41ac-9af8-83330e9fc0c8.jpg",     // реальное фото станка в цехе
};

const RED = "C0102C";
const WHITE = "FFFFFF";
const LIGHT_GRAY = "F4F4F4";
const DARK = "1A1A1A";
const MEDIUM_GRAY = "E8E8E8";

// Slide data (for preview)
const slides = [
  { id: 1, title: "GiperLaser Extra", subtitle: "Разработка инновационных машин термической резки металлов", type: "cover" },
  {
    id: 2, title: "О компании", type: "points-img", img: IMGS.realMachine, imgCaption: "Производство, Красноярск",
    points: ["ООО «Гиперплазма», Красноярск — ОКВЭД 28.41.1", "Резидент Сколково и МТК с 2025 года", "14 специалистов, 150+ реализованных проектов", "Патенты, ПО, Реестр российской промышленной продукции", "Производство: Красноярск и пос. Солонцы"],
  },
  {
    id: 3, title: "Проблема и цель", type: "points",
    points: ["Рынок зависит от оборудования Bystronic, Trumpf, Mazak, Amada", "Поставки из ЕС и КНР ограничены — сервис недоступен", "Цель: создать российский станок с мировыми характеристиками", "Адаптация к российским условиям и оперативному сервису"],
  },
  {
    id: 4, title: "Продукт: GiperLaser Extra", type: "points-img", img: IMGS.machine3d, imgCaption: "GiperLaser Extra — 3D рендер",
    points: ["Лазерный станок с ЧПУ для резки металла", "Скорость резки до 80 м/мин, ускорение 0,8G", "Точность позиционирования ±0,05 мм на 4 м", "Резка до 100 мм: нержавейка, алюминий и др.", "Лазер 12–60 кВт, рабочее поле 6×2 м"],
  },
  {
    id: 5, title: "Конструкция: ключевые узлы", type: "three-imgs",
    images: [
      { src: IMGS.portalBeam, label: "Портальная балка" },
      { src: IMGS.machineFrame, label: "Сварная рама станка" },
      { src: IMGS.zAxis, label: "Узел оси Z / суппорт" },
    ],
  },
  {
    id: 6, title: "Методология НИОКР", type: "cols-imgs",
    cols: [
      { label: "Анализ", items: ["ТРИЗ", "Функциональный анализ", "Морфологический анализ"], img: IMGS.metalParts },
      { label: "Оптимизация", items: ["Топологическая", "Генеративная", "FEA-моделирование"], img: IMGS.blackParts },
      { label: "Разработка", items: ["Прототипирование", "Испытания", "Адаптивное управление"], img: IMGS.brassClamp },
    ],
  },
  {
    id: 7, title: "Рынок и конкуренты", type: "stats",
    stats: [{ value: "$350–400M", label: "Объём рынка РФ" }, { value: "8–10%", label: "CAGR до 2029 г." }, { value: "3 200–3 500", label: "Установок/год к 2029" }, { value: "60%", label: "Целевая доля отечественных" }],
    note: "Замещаем: Bystronic · Trumpf · Mazak · Amada · Bodor · Senfeng",
  },
  {
    id: 8, title: "Целевые клиенты", type: "points-img", img: IMGS.techDiagram, imgCaption: "Схема режущей головы GiperLaser",
    points: ["Русал, СУЭК, ГМК Норильский никель", "ГК Росатом, Транснефть, Роснефть, Газпром", "Северсталь, Звёздочка", "Заводы металлоконструкций и машиностроения", "Рынки: Россия, Казахстан, Беларусь"],
  },
  {
    id: 9, title: "Финансовая модель", type: "stats",
    stats: [{ value: "от 16M ₽", label: "Минимальная цена станка" }, { value: "124,8M ₽", label: "Выручка 2026–2027" }, { value: "37,4M ₽", label: "Чистая прибыль" }, { value: "125%", label: "ROI" }],
    note: "Окупаемость от 10 месяцев · Точка безубыточности — 10 станков/год",
  },
  {
    id: 10, title: "Следующий шаг", subtitle: "Приглашаем к сотрудничеству", type: "final",
    points: ["Инвестиции в НИОКР: 30 млн руб.", "Статус резидента Сколково — налоговые льготы", "Команда с опытом 150+ проектов в станкостроении", "Первые поставки — 2026 год"],
    contact: "ООО «Гиперплазма» · Красноярск · giplasma.ru",
  },
];

async function generatePptx() {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  pptx.title = "GiperLaser Extra — Инновационные машины лазерной резки";

  const addDecorLine = (slide: PptxGenJS.Slide) => {
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.22, h: 5.63, fill: { color: RED } });
  };

  const addFooter = (slide: PptxGenJS.Slide, num: number) => {
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 5.2, w: 10, h: 0.43, fill: { color: MEDIUM_GRAY } });
    slide.addText("ООО «Гиперплазма» · GiperLaser Extra · НИОКР 2025–2027", { x: 0.35, y: 5.23, w: 8.5, h: 0.28, fontSize: 9, color: "666666", fontFace: "IBM Plex Sans" });
    slide.addText(`${num} / 10`, { x: 8.8, y: 5.23, w: 1, h: 0.28, fontSize: 9, color: RED, fontFace: "Montserrat", bold: true, align: "right" });
  };

  const addHeader = (slide: PptxGenJS.Slide, title: string) => {
    slide.addShape(pptx.ShapeType.rect, { x: 0.22, y: 0, w: 9.78, h: 1.1, fill: { color: RED } });
    slide.addText(title, { x: 0.42, y: 0.15, w: 9.2, h: 0.75, fontSize: 26, bold: true, color: WHITE, fontFace: "Montserrat" });
  };

  // --- Slide 1: Cover ---
  const s1 = pptx.addSlide();
  s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 5.63, fill: { color: DARK } });
  s1.addImage({ path: IMGS.machine3d, x: 3.8, y: 0.3, w: 6.0, h: 5.0 });
  s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 5.2, h: 5.63, fill: { color: DARK } });
  s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.18, h: 5.63, fill: { color: RED } });
  s1.addText("GiperLaser", { x: 0.35, y: 1.0, w: 4.5, h: 0.8, fontSize: 38, bold: true, color: WHITE, fontFace: "Montserrat" });
  s1.addText("Extra", { x: 0.35, y: 1.75, w: 4.5, h: 0.7, fontSize: 38, bold: true, color: RED, fontFace: "Montserrat" });
  s1.addShape(pptx.ShapeType.rect, { x: 0.35, y: 2.55, w: 3.5, h: 0.04, fill: { color: RED } });
  s1.addText("Разработка инновационных\nмашин термической резки металлов", { x: 0.35, y: 2.7, w: 4.5, h: 0.9, fontSize: 13, color: "CCCCCC", fontFace: "IBM Plex Sans" });
  s1.addText("Импортозамещение  ·  НИОКР  ·  2025–2027", { x: 0.35, y: 3.8, w: 4.5, h: 0.4, fontSize: 10, color: RED, fontFace: "Montserrat", bold: true });
  s1.addText("ООО «Гиперплазма»  ·  Красноярск  ·  Резидент Сколково", { x: 0.35, y: 4.9, w: 4.5, h: 0.3, fontSize: 9, color: "888888", fontFace: "IBM Plex Sans" });

  // --- Slide 2: О компании (points + real photo) ---
  const s2 = pptx.addSlide();
  s2.background = { color: WHITE };
  addDecorLine(s2); addFooter(s2, 2); addHeader(s2, "О компании");
  s2.addImage({ path: IMGS.realMachine, x: 5.8, y: 1.2, w: 3.8, h: 2.8, rounding: false });
  s2.addShape(pptx.ShapeType.rect, { x: 5.8, y: 3.9, w: 3.8, h: 0.3, fill: { color: RED } });
  s2.addText("Производство, Красноярск", { x: 5.85, y: 3.92, w: 3.7, h: 0.26, fontSize: 9, color: WHITE, fontFace: "IBM Plex Sans", italic: true });
  const s2pts = ["ООО «Гиперплазма», Красноярск — ОКВЭД 28.41.1", "Резидент Сколково и МТК с 2025 года", "14 специалистов, 150+ реализованных проектов", "Патенты, ПО, Реестр российской промышленной продукции", "Производство: Красноярск и пос. Солонцы"];
  s2pts.forEach((pt, i) => {
    const y = 1.2 + i * 0.68;
    s2.addShape(pptx.ShapeType.rect, { x: 0.42, y: y + 0.12, w: 0.18, h: 0.18, fill: { color: RED } });
    s2.addText(pt, { x: 0.72, y, w: 4.9, h: 0.55, fontSize: 13.5, color: DARK, fontFace: "IBM Plex Sans" });
  });

  // --- Slide 3: Проблема и цель ---
  const s3 = pptx.addSlide();
  s3.background = { color: WHITE };
  addDecorLine(s3); addFooter(s3, 3); addHeader(s3, "Проблема и цель");
  const s3pts = ["Рынок зависит от оборудования Bystronic, Trumpf, Mazak, Amada", "Поставки из ЕС и КНР ограничены — сервис недоступен", "Цель: создать российский станок с мировыми характеристиками", "Адаптация к российским условиям и оперативному сервису"];
  s3pts.forEach((pt, i) => {
    const y = 1.4 + i * 0.85;
    s3.addShape(pptx.ShapeType.rect, { x: 0.42, y: y + 0.15, w: 0.22, h: 0.22, fill: { color: RED } });
    s3.addText(pt, { x: 0.78, y, w: 8.8, h: 0.65, fontSize: 15, color: DARK, fontFace: "IBM Plex Sans" });
  });

  // --- Slide 4: Продукт (points + 3D render) ---
  const s4 = pptx.addSlide();
  s4.background = { color: WHITE };
  addDecorLine(s4); addFooter(s4, 4); addHeader(s4, "Продукт: GiperLaser Extra");
  s4.addImage({ path: IMGS.machine3d, x: 5.5, y: 1.1, w: 4.1, h: 3.1 });
  s4.addShape(pptx.ShapeType.rect, { x: 5.5, y: 4.1, w: 4.1, h: 0.28, fill: { color: LIGHT_GRAY } });
  s4.addText("GiperLaser Extra — 3D рендер", { x: 5.55, y: 4.12, w: 4.0, h: 0.24, fontSize: 8.5, color: "888888", fontFace: "IBM Plex Sans", italic: true });
  const s4pts = ["Лазерный станок с ЧПУ для резки металла", "Скорость резки до 80 м/мин, ускорение 0,8G", "Точность позиционирования ±0,05 мм на 4 м", "Резка до 100 мм: нержавейка, алюминий и др.", "Лазер 12–60 кВт, рабочее поле 6×2 м"];
  s4pts.forEach((pt, i) => {
    const y = 1.2 + i * 0.68;
    s4.addShape(pptx.ShapeType.rect, { x: 0.42, y: y + 0.12, w: 0.18, h: 0.18, fill: { color: RED } });
    s4.addText(pt, { x: 0.72, y, w: 4.6, h: 0.55, fontSize: 13.5, color: DARK, fontFace: "IBM Plex Sans" });
  });

  // --- Slide 5: Конструкция — 3 картинки ---
  const s5 = pptx.addSlide();
  s5.background = { color: WHITE };
  addDecorLine(s5); addFooter(s5, 5); addHeader(s5, "Конструкция: ключевые узлы");
  const imgData5 = [
    { path: IMGS.portalBeam, label: "Портальная балка" },
    { path: IMGS.machineFrame, label: "Сварная рама станка" },
    { path: IMGS.zAxis, label: "Узел оси Z / суппорт" },
  ];
  imgData5.forEach((img, i) => {
    const x = 0.35 + i * 3.2;
    s5.addImage({ path: img.path, x, y: 1.2, w: 3.0, h: 2.9 });
    s5.addShape(pptx.ShapeType.rect, { x, y: 4.1, w: 3.0, h: 0.38, fill: { color: RED } });
    s5.addText(img.label, { x: x + 0.05, y: 4.13, w: 2.9, h: 0.32, fontSize: 11, color: WHITE, fontFace: "Montserrat", bold: true, align: "center" });
  });

  // --- Slide 6: НИОКР — колонки с картинками ---
  const s6 = pptx.addSlide();
  s6.background = { color: WHITE };
  addDecorLine(s6); addFooter(s6, 6); addHeader(s6, "Методология НИОКР");
  const cols6 = [
    { label: "Анализ", items: ["ТРИЗ", "Функциональный анализ", "Морфологический анализ"], img: IMGS.metalParts },
    { label: "Оптимизация", items: ["Топологическая", "Генеративная", "FEA-моделирование"], img: IMGS.blackParts },
    { label: "Разработка", items: ["Прототипирование", "Испытания", "Адаптивное управление"], img: IMGS.brassClamp },
  ];
  cols6.forEach((col, ci) => {
    const x = 0.35 + ci * 3.2;
    s6.addShape(pptx.ShapeType.rect, { x, y: 1.2, w: 3.0, h: 0.5, fill: { color: RED } });
    s6.addText(col.label, { x: x + 0.1, y: 1.25, w: 2.8, h: 0.4, fontSize: 13, bold: true, color: WHITE, fontFace: "Montserrat" });
    s6.addImage({ path: col.img, x, y: 1.7, w: 3.0, h: 1.8 });
    col.items.forEach((item, ii) => {
      s6.addText(`• ${item}`, { x: x + 0.1, y: 3.55 + ii * 0.45, w: 2.8, h: 0.4, fontSize: 12, color: DARK, fontFace: "IBM Plex Sans" });
    });
  });

  // --- Slide 7: Рынок ---
  const s7 = pptx.addSlide();
  s7.background = { color: WHITE };
  addDecorLine(s7); addFooter(s7, 7); addHeader(s7, "Рынок и конкуренты");
  const stats7 = [{ value: "$350–400M", label: "Объём рынка РФ" }, { value: "8–10%", label: "CAGR до 2029 г." }, { value: "3 200–3 500", label: "Установок/год к 2029" }, { value: "60%", label: "Целевая доля отечественных" }];
  stats7.forEach((st, si) => {
    const x = 0.42 + (si % 2) * 4.65;
    const y = si < 2 ? 1.25 : 3.1;
    s7.addShape(pptx.ShapeType.rect, { x, y, w: 4.35, h: 1.6, fill: { color: LIGHT_GRAY } });
    s7.addShape(pptx.ShapeType.rect, { x, y, w: 0.1, h: 1.6, fill: { color: RED } });
    s7.addText(st.value, { x: x + 0.2, y: y + 0.15, w: 4.0, h: 0.7, fontSize: 26, bold: true, color: RED, fontFace: "Montserrat" });
    s7.addText(st.label, { x: x + 0.2, y: y + 0.85, w: 4.0, h: 0.55, fontSize: 13, color: DARK, fontFace: "IBM Plex Sans" });
  });
  s7.addShape(pptx.ShapeType.rect, { x: 0.42, y: 4.8, w: 9.2, h: 0.32, fill: { color: MEDIUM_GRAY } });
  s7.addText("Замещаем: Bystronic · Trumpf · Mazak · Amada · Bodor · Senfeng", { x: 0.6, y: 4.82, w: 9.0, h: 0.28, fontSize: 10.5, color: "444444", fontFace: "IBM Plex Sans", italic: true });

  // --- Slide 8: Клиенты + схема головы ---
  const s8 = pptx.addSlide();
  s8.background = { color: WHITE };
  addDecorLine(s8); addFooter(s8, 8); addHeader(s8, "Целевые клиенты");
  s8.addImage({ path: IMGS.techDiagram, x: 5.9, y: 1.15, w: 3.6, h: 3.3 });
  s8.addShape(pptx.ShapeType.rect, { x: 5.9, y: 4.4, w: 3.6, h: 0.3, fill: { color: RED } });
  s8.addText("Схема режущей головы GiperLaser", { x: 5.95, y: 4.42, w: 3.5, h: 0.26, fontSize: 9, color: WHITE, fontFace: "IBM Plex Sans", italic: true });
  const s8pts = ["Русал, СУЭК, ГМК Норильский никель", "ГК Росатом, Транснефть, Роснефть, Газпром", "Северсталь, Звёздочка", "Заводы металлоконструкций и машиностроения", "Рынки: Россия, Казахстан, Беларусь"];
  s8pts.forEach((pt, i) => {
    const y = 1.2 + i * 0.68;
    s8.addShape(pptx.ShapeType.rect, { x: 0.42, y: y + 0.12, w: 0.18, h: 0.18, fill: { color: RED } });
    s8.addText(pt, { x: 0.72, y, w: 5.0, h: 0.55, fontSize: 13.5, color: DARK, fontFace: "IBM Plex Sans" });
  });

  // --- Slide 9: Финансы ---
  const s9 = pptx.addSlide();
  s9.background = { color: WHITE };
  addDecorLine(s9); addFooter(s9, 9); addHeader(s9, "Финансовая модель");
  const stats9 = [{ value: "от 16M ₽", label: "Минимальная цена станка" }, { value: "124,8M ₽", label: "Выручка 2026–2027" }, { value: "37,4M ₽", label: "Чистая прибыль" }, { value: "125%", label: "ROI" }];
  stats9.forEach((st, si) => {
    const x = 0.42 + (si % 2) * 4.65;
    const y = si < 2 ? 1.25 : 3.1;
    s9.addShape(pptx.ShapeType.rect, { x, y, w: 4.35, h: 1.6, fill: { color: LIGHT_GRAY } });
    s9.addShape(pptx.ShapeType.rect, { x, y, w: 0.1, h: 1.6, fill: { color: RED } });
    s9.addText(st.value, { x: x + 0.2, y: y + 0.15, w: 4.0, h: 0.7, fontSize: 26, bold: true, color: RED, fontFace: "Montserrat" });
    s9.addText(st.label, { x: x + 0.2, y: y + 0.85, w: 4.0, h: 0.55, fontSize: 13, color: DARK, fontFace: "IBM Plex Sans" });
  });
  s9.addShape(pptx.ShapeType.rect, { x: 0.42, y: 4.8, w: 9.2, h: 0.32, fill: { color: MEDIUM_GRAY } });
  s9.addText("Окупаемость от 10 месяцев · Точка безубыточности — 10 станков/год", { x: 0.6, y: 4.82, w: 9.0, h: 0.28, fontSize: 10.5, color: "444444", fontFace: "IBM Plex Sans", italic: true });

  // --- Slide 10: Final ---
  const s10 = pptx.addSlide();
  s10.background = { color: WHITE };
  addDecorLine(s10); addFooter(s10, 10); addHeader(s10, "Следующий шаг");
  s10.addText("Приглашаем к сотрудничеству", { x: 0.42, y: 1.15, w: 9.2, h: 0.45, fontSize: 16, color: "444444", fontFace: "IBM Plex Sans" });
  const s10pts = ["Инвестиции в НИОКР: 30 млн руб.", "Статус резидента Сколково — налоговые льготы", "Команда с опытом 150+ проектов в станкостроении", "Первые поставки — 2026 год"];
  s10pts.forEach((pt, i) => {
    const y = 1.75 + i * 0.68;
    s10.addShape(pptx.ShapeType.rect, { x: 0.42, y: y + 0.12, w: 0.18, h: 0.18, fill: { color: RED } });
    s10.addText(pt, { x: 0.72, y, w: 8.8, h: 0.55, fontSize: 14.5, color: DARK, fontFace: "IBM Plex Sans" });
  });
  s10.addShape(pptx.ShapeType.rect, { x: 0.42, y: 4.6, w: 9.2, h: 0.45, fill: { color: RED } });
  s10.addText("ООО «Гиперплазма» · Красноярск · giplasma.ru", { x: 0.6, y: 4.63, w: 9.0, h: 0.38, fontSize: 12, color: WHITE, fontFace: "Montserrat", bold: true, align: "center" });

  await pptx.writeFile({ fileName: "GiperLaser_Extra_Presentation.pptx" });
}

// ---- Preview Component ----
type SlideAny = { id: number; title: string; type: string; subtitle?: string; points?: string[]; img?: string; images?: { src: string; label: string }[]; cols?: { label: string; items: string[]; img: string }[]; stats?: { value: string; label: string }[]; note?: string; contact?: string };

const SlidePreview = ({ slide, index }: { slide: SlideAny; index: number }) => {
  const s = slide;
  return (
    <div className="relative bg-white rounded-sm overflow-hidden shadow-md hover:shadow-xl transition-all duration-300 hover:-translate-y-1">
      <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-red-700 z-10" />
      <div className="absolute top-2 right-2 bg-red-700 text-white text-[9px] font-montserrat font-bold w-5 h-5 rounded-full flex items-center justify-center z-20">{index + 1}</div>

      {slide.type === "cover" ? (
        <div className="aspect-[16/9] bg-gray-900 relative overflow-hidden">
          <img src={IMGS.machine3d} alt="" className="absolute right-0 top-0 w-3/5 h-full object-contain" />
          <div className="absolute inset-0 bg-gradient-to-r from-gray-900 via-gray-900/85 to-transparent" />
          <div className="absolute left-5 top-1/2 -translate-y-1/2">
            <div className="text-white font-black font-montserrat text-xl leading-tight">GiperLaser<br /><span className="text-red-500">Extra</span></div>
            <div className="w-10 h-0.5 bg-red-600 my-1.5" />
            <div className="text-gray-300 text-[8px] font-ibm leading-tight">Разработка инновационных<br />машин термической резки</div>
            <div className="text-red-400 text-[7px] font-montserrat font-bold mt-1.5 tracking-wide">НИОКР · 2025–2027</div>
          </div>
        </div>
      ) : (
        <div className="aspect-[16/9] bg-white relative overflow-hidden">
          <div className="absolute top-0 left-1.5 right-0 h-[20%] bg-red-700 flex items-center px-2">
            <span className="text-white font-black font-montserrat text-[10px] leading-tight truncate">{slide.title}</span>
          </div>
          <div className="absolute top-[22%] left-2.5 right-1.5 bottom-[10%] overflow-hidden">

            {(slide.type === "points" || slide.type === "points-img") && (
              <div className="flex gap-1.5 h-full">
                <div className="flex-1 space-y-0.5">
                  {s.points?.map((pt: string, i: number) => (
                    <div key={i} className="flex items-start gap-1">
                      <div className="w-1 h-1 bg-red-600 mt-1 shrink-0" />
                      <span className="text-gray-800 text-[7px] font-ibm leading-tight">{pt}</span>
                    </div>
                  ))}
                </div>
                {s.img && (
                  <div className="w-[42%] shrink-0">
                    <img src={s.img} alt="" className="w-full h-full object-cover rounded-sm" />
                  </div>
                )}
              </div>
            )}

            {slide.type === "three-imgs" && (
              <div className="flex gap-1 h-full">
                {s.images?.map((img: { src: string; label: string }, i: number) => (
                  <div key={i} className="flex-1 flex flex-col">
                    <img src={img.src} alt="" className="flex-1 w-full object-cover" />
                    <div className="bg-red-600 text-white text-[6px] font-montserrat font-bold text-center py-0.5 px-0.5 leading-tight">{img.label}</div>
                  </div>
                ))}
              </div>
            )}

            {slide.type === "cols-imgs" && (
              <div className="flex gap-1 h-full">
                {s.cols?.map((col: { label: string; items: string[]; img: string }, ci: number) => (
                  <div key={ci} className="flex-1 flex flex-col overflow-hidden">
                    <div className="bg-red-600 px-1 py-0.5"><span className="text-white text-[7px] font-montserrat font-bold">{col.label}</span></div>
                    <img src={col.img} alt="" className="w-full h-[45%] object-cover" />
                    <div className="flex-1 p-0.5 space-y-0.5 bg-gray-50">
                      {col.items.map((item, ii) => <div key={ii} className="text-gray-700 text-[6px] font-ibm">• {item}</div>)}
                    </div>
                  </div>
                ))}
              </div>
            )}

            {slide.type === "stats" && (
              <div className="grid grid-cols-2 gap-1 h-full">
                {s.stats?.map((st: { value: string; label: string }, si: number) => (
                  <div key={si} className="bg-gray-50 border-l-2 border-red-600 px-1 py-0.5 flex flex-col justify-center">
                    <div className="text-red-700 font-black font-montserrat text-[10px] leading-tight">{st.value}</div>
                    <div className="text-gray-600 text-[6px] font-ibm leading-tight mt-0.5">{st.label}</div>
                  </div>
                ))}
              </div>
            )}

            {slide.type === "final" && (
              <div className="space-y-0.5">
                {s.subtitle && <div className="text-gray-500 text-[7px] font-ibm mb-1">{s.subtitle}</div>}
                {s.points?.map((pt: string, i: number) => (
                  <div key={i} className="flex items-start gap-1">
                    <div className="w-1 h-1 bg-red-600 mt-1 shrink-0" />
                    <span className="text-gray-800 text-[7px] font-ibm leading-tight">{pt}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
          <div className="absolute bottom-0 left-1.5 right-0 h-[10%] bg-gray-100 flex items-center justify-between px-2">
            <span className="text-gray-400 text-[6px] font-ibm truncate">ООО «Гиперплазма» · GiperLaser Extra</span>
            <span className="text-red-600 text-[7px] font-montserrat font-bold">{index + 1}/10</span>
          </div>
        </div>
      )}
    </div>
  );
};

const Index = () => {
  const [generating, setGenerating] = useState(false);

  const handleDownload = async () => {
    setGenerating(true);
    try { await generatePptx(); } finally { setGenerating(false); }
  };

  return (
    <div className="min-h-screen bg-[#F4F4F4] font-ibm">
      {/* Top bar */}
      <div className="bg-[#1A1A1A] border-b-4 border-red-700">
        <div className="max-w-7xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="w-1 h-8 bg-red-700" />
            <div>
              <div className="text-white font-black font-montserrat text-xl tracking-tight leading-none">
                GiperLaser <span className="text-red-500">Extra</span>
              </div>
              <div className="text-gray-400 text-xs font-ibm mt-0.5">Инновационные машины лазерной резки металлов</div>
            </div>
          </div>
          <button onClick={handleDownload} disabled={generating}
            className="flex items-center gap-2.5 bg-red-700 hover:bg-red-800 disabled:opacity-70 text-white font-montserrat font-bold text-sm px-6 py-3 transition-all duration-200 hover:shadow-lg hover:shadow-red-900/40 active:scale-95">
            <Icon name={generating ? "Loader2" : "Download"} size={16} className={generating ? "animate-spin" : ""} />
            {generating ? "Генерация..." : "Скачать PPTX"}
          </button>
        </div>
      </div>

      {/* Info bar */}
      <div className="bg-white border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-6 py-6 flex items-center justify-between">
          <div>
            <h1 className="text-3xl font-black font-montserrat text-gray-900">Презентация проекта</h1>
            <p className="text-gray-500 font-ibm mt-1 text-sm">10 слайдов · PowerPoint · Строгий деловой стиль · С фотографиями</p>
          </div>
          <div className="flex gap-6 text-center">
            {[{ val: "10", label: "слайдов" }, { val: "8", label: "фото" }, { val: "PPTX", label: "формат" }].map(item => (
              <div key={item.label}>
                <div className="text-2xl font-black font-montserrat text-red-700">{item.val}</div>
                <div className="text-xs text-gray-400 font-ibm">{item.label}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Slides grid */}
      <div className="max-w-7xl mx-auto px-6 py-10">
        <div className="flex items-center gap-3 mb-6">
          <div className="w-1 h-5 bg-red-700" />
          <h2 className="text-lg font-bold font-montserrat text-gray-800">Предпросмотр слайдов</h2>
        </div>
        <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5 gap-4">
          {slides.map((slide, i) => <SlidePreview key={slide.id} slide={slide} index={i} />)}
        </div>
      </div>

      {/* Bottom CTA */}
      <div className="border-t border-gray-200 bg-white">
        <div className="max-w-7xl mx-auto px-6 py-8 flex flex-col sm:flex-row items-center justify-between gap-4">
          <div className="text-sm text-gray-500 font-ibm">
            Файл <strong>.pptx</strong> открывается в PowerPoint, Keynote, Google Slides
          </div>
          <button onClick={handleDownload} disabled={generating}
            className="flex items-center gap-2.5 bg-red-700 hover:bg-red-800 disabled:opacity-70 text-white font-montserrat font-bold text-sm px-8 py-3.5 transition-all duration-200 hover:shadow-lg hover:shadow-red-900/40 active:scale-95">
            <Icon name={generating ? "Loader2" : "FileDown"} size={16} className={generating ? "animate-spin" : ""} />
            {generating ? "Генерация файла..." : "Скачать презентацию"}
          </button>
        </div>
      </div>
    </div>
  );
};

export default Index;