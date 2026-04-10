import { useState } from "react";
import Icon from "@/components/ui/icon";
import PptxGenJS from "pptxgenjs";

const LASER_IMAGE = "https://cdn.poehali.dev/projects/edca03ca-d59d-4b05-889c-21f49114664a/files/01e369cc-f9a3-4c2e-a810-fc342bc061d3.jpg";

const RED = "C0102C";
const WHITE = "FFFFFF";
const LIGHT_GRAY = "F4F4F4";
const DARK = "1A1A1A";
const MEDIUM_GRAY = "E8E8E8";

const slides = [
  {
    id: 1,
    title: "GiperLaser Extra",
    subtitle: "Разработка инновационных машин термической резки металлов",
    tag: "Импортозамещение · НИОКР · 2025–2027",
    type: "cover",
  },
  {
    id: 2,
    title: "О компании",
    points: [
      "ООО «Гиперплазма», Красноярск — ОКВЭД 28.41.1",
      "Резидент Сколково и МТК с 2025 года",
      "14 специалистов, 150+ реализованных проектов",
      "Собственные патенты и ПО, продукция в Реестре РПП",
      "Производство: Красноярск и пос. Солонцы",
    ],
    icon: "Building2",
    type: "points",
  },
  {
    id: 3,
    title: "Проблема и цель",
    points: [
      "Рынок зависит от оборудования Bystronic, Trumpf, Mazak, Amada",
      "Поставки из ЕС и КНР ограничены — сервис недоступен",
      "Цель: создать российский станок с мировыми характеристиками",
      "Адаптация к российским условиям и оперативному сервису",
    ],
    icon: "Target",
    type: "points",
  },
  {
    id: 4,
    title: "Продукт: GiperLaser Extra",
    points: [
      "Лазерный станок с ЧПУ для резки металла",
      "Скорость резки до 80 м/мин, ускорение 0,8G",
      "Точность позиционирования ±0,05 мм на 4 м",
      "Резка до 100 мм: нержавейка, алюминий и др.",
      "Лазер 12–60 кВт, рабочее поле 6×2 м",
    ],
    icon: "Zap",
    type: "points",
  },
  {
    id: 5,
    title: "Технические инновации",
    points: [
      "Модульная архитектура конструкции портала",
      "3D‑суппорт и адаптивная система управления",
      "Топологическая и генеративная оптимизация (FEA)",
      "Системы выборки люфта и поиска заготовки",
      "Защита режущей головы + интеграция сенсоров",
    ],
    icon: "Cpu",
    type: "points",
  },
  {
    id: 6,
    title: "Методология НИОКР",
    cols: [
      { label: "Анализ", items: ["ТРИЗ", "Функциональный анализ", "Морфологический анализ"] },
      { label: "Оптимизация", items: ["Топологическая", "Генеративная", "FEA-моделирование"] },
      { label: "Разработка", items: ["Прототипирование", "Испытания", "Адаптивное управление"] },
    ],
    type: "cols",
  },
  {
    id: 7,
    title: "Рынок и конкуренты",
    stats: [
      { value: "$350–400M", label: "Объём рынка РФ" },
      { value: "8–10%", label: "CAGR до 2029 г." },
      { value: "3 200–3 500", label: "Установок/год к 2029" },
      { value: "60%", label: "Целевая доля отечественных" },
    ],
    note: "Замещаем: Bystronic · Trumpf · Mazak · Amada · Bodor · Senfeng",
    type: "stats",
  },
  {
    id: 8,
    title: "Целевые клиенты",
    points: [
      "Русал, СУЭК, ГМК Норильский никель",
      "ГК Росатом, Транснефть, Роснефть, Газпром",
      "Северсталь, Звёздочка",
      "Заводы металлоконструкций и машиностроения",
      "Рынки: Россия, Казахстан, Беларусь",
    ],
    icon: "Users",
    type: "points",
  },
  {
    id: 9,
    title: "Финансовая модель",
    stats: [
      { value: "от 16M ₽", label: "Минимальная цена станка" },
      { value: "124,8M ₽", label: "Выручка 2026–2027" },
      { value: "37,4M ₽", label: "Чистая прибыль" },
      { value: "125%", label: "ROI" },
    ],
    note: "Окупаемость от 10 месяцев · Точка безубыточности — 10 станков/год",
    type: "stats",
  },
  {
    id: 10,
    title: "Следующий шаг",
    subtitle: "Приглашаем к сотрудничеству",
    points: [
      "Инвестиции в НИОКР: 30 млн руб.",
      "Статус резидента Сколково — налоговые льготы",
      "Команда с опытом 150+ проектов в станкостроении",
      "Первые поставки — 2026 год",
    ],
    contact: "ООО «Гиперплазма» · Красноярск · giplasma.ru",
    type: "final",
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
    slide.addText("ООО «Гиперплазма» · GiperLaser Extra · НИОКР 2025–2027", {
      x: 0.35, y: 5.23, w: 8.5, h: 0.28,
      fontSize: 9, color: "666666", fontFace: "IBM Plex Sans",
    });
    slide.addText(`${num} / 10`, {
      x: 8.8, y: 5.23, w: 1, h: 0.28,
      fontSize: 9, color: RED, fontFace: "Montserrat", bold: true, align: "right",
    });
  };

  // Slide 1 — Cover
  const s1 = pptx.addSlide();
  s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 5.63, fill: { color: DARK } });
  s1.addImage({ path: LASER_IMAGE, x: 4.5, y: 0, w: 5.5, h: 5.63 });
  s1.addShape(pptx.ShapeType.rect, { x: 4.5, y: 0, w: 5.5, h: 5.63, fill: { color: DARK }, transparency: 55 });
  s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 4.8, h: 5.63, fill: { color: DARK } });
  s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.18, h: 5.63, fill: { color: RED } });
  s1.addText("GiperLaser", { x: 0.35, y: 1.0, w: 4.2, h: 0.8, fontSize: 38, bold: true, color: WHITE, fontFace: "Montserrat" });
  s1.addText("Extra", { x: 0.35, y: 1.75, w: 4.2, h: 0.7, fontSize: 38, bold: true, color: RED, fontFace: "Montserrat" });
  s1.addShape(pptx.ShapeType.rect, { x: 0.35, y: 2.55, w: 3.5, h: 0.04, fill: { color: RED } });
  s1.addText("Разработка инновационных\nмашин термической резки металлов", {
    x: 0.35, y: 2.7, w: 4.1, h: 0.9, fontSize: 13, color: "CCCCCC", fontFace: "IBM Plex Sans",
  });
  s1.addText("Импортозамещение  ·  НИОКР  ·  2025–2027", {
    x: 0.35, y: 3.8, w: 4.1, h: 0.4, fontSize: 10, color: RED, fontFace: "Montserrat", bold: true,
  });
  s1.addText("ООО «Гиперплазма»  ·  Красноярск  ·  Резидент Сколково", {
    x: 0.35, y: 4.9, w: 4.1, h: 0.3, fontSize: 9, color: "888888", fontFace: "IBM Plex Sans",
  });

  // Slides 2–10
  const contentSlides = slides.slice(1);
  for (let i = 0; i < contentSlides.length; i++) {
    const data = contentSlides[i];
    const slide = pptx.addSlide();
    slide.background = { color: WHITE };
    addDecorLine(slide);
    addFooter(slide, i + 2);

    // Red header bar
    slide.addShape(pptx.ShapeType.rect, { x: 0.22, y: 0, w: 9.78, h: 1.1, fill: { color: RED } });
    slide.addText(data.title, {
      x: 0.42, y: 0.15, w: 9.2, h: 0.75,
      fontSize: 26, bold: true, color: WHITE, fontFace: "Montserrat",
    });

    if (data.type === "points" && data.points) {
      data.points.forEach((pt, idx) => {
        const y = 1.3 + idx * 0.68;
        slide.addShape(pptx.ShapeType.rect, { x: 0.42, y: y + 0.12, w: 0.18, h: 0.18, fill: { color: RED } });
        slide.addText(pt, {
          x: 0.72, y, w: 8.8, h: 0.55,
          fontSize: 14.5, color: DARK, fontFace: "IBM Plex Sans",
        });
      });
    }

    if (data.type === "cols" && data.cols) {
      data.cols.forEach((col, ci) => {
        const x = 0.42 + ci * 3.15;
        slide.addShape(pptx.ShapeType.rect, { x, y: 1.25, w: 2.9, h: 3.7, fill: { color: LIGHT_GRAY } });
        slide.addShape(pptx.ShapeType.rect, { x, y: 1.25, w: 2.9, h: 0.5, fill: { color: RED } });
        slide.addText(col.label, {
          x: x + 0.1, y: 1.3, w: 2.7, h: 0.4,
          fontSize: 13, bold: true, color: WHITE, fontFace: "Montserrat",
        });
        col.items.forEach((item, ii) => {
          slide.addText(`• ${item}`, {
            x: x + 0.15, y: 1.9 + ii * 0.65, w: 2.6, h: 0.55,
            fontSize: 13, color: DARK, fontFace: "IBM Plex Sans",
          });
        });
      });
    }

    if (data.type === "stats" && data.stats) {
      data.stats.forEach((st, si) => {
        const x = 0.42 + (si % 2) * 4.65;
        const y = si < 2 ? 1.25 : 3.1;
        slide.addShape(pptx.ShapeType.rect, { x, y, w: 4.35, h: 1.6, fill: { color: LIGHT_GRAY } });
        slide.addShape(pptx.ShapeType.rect, { x, y, w: 0.1, h: 1.6, fill: { color: RED } });
        slide.addText(st.value, {
          x: x + 0.2, y: y + 0.15, w: 4.0, h: 0.7,
          fontSize: 26, bold: true, color: RED, fontFace: "Montserrat",
        });
        slide.addText(st.label, {
          x: x + 0.2, y: y + 0.85, w: 4.0, h: 0.55,
          fontSize: 13, color: DARK, fontFace: "IBM Plex Sans",
        });
      });
      if (data.note) {
        slide.addShape(pptx.ShapeType.rect, { x: 0.42, y: 4.8, w: 9.2, h: 0.32, fill: { color: MEDIUM_GRAY } });
        slide.addText(data.note, {
          x: 0.6, y: 4.82, w: 9.0, h: 0.28,
          fontSize: 10.5, color: "444444", fontFace: "IBM Plex Sans", italic: true,
        });
      }
    }

    if (data.type === "final") {
      if (data.subtitle) {
        slide.addText(data.subtitle, {
          x: 0.42, y: 1.15, w: 9.2, h: 0.45,
          fontSize: 16, color: "444444", fontFace: "IBM Plex Sans",
        });
      }
      if (data.points) {
        data.points.forEach((pt, idx) => {
          const y = 1.75 + idx * 0.68;
          slide.addShape(pptx.ShapeType.rect, { x: 0.42, y: y + 0.12, w: 0.18, h: 0.18, fill: { color: RED } });
          slide.addText(pt, {
            x: 0.72, y, w: 8.8, h: 0.55,
            fontSize: 14.5, color: DARK, fontFace: "IBM Plex Sans",
          });
        });
      }
      if (data.contact) {
        slide.addShape(pptx.ShapeType.rect, { x: 0.42, y: 4.6, w: 9.2, h: 0.45, fill: { color: RED } });
        slide.addText(data.contact, {
          x: 0.6, y: 4.63, w: 9.0, h: 0.38,
          fontSize: 12, color: WHITE, fontFace: "Montserrat", bold: true, align: "center",
        });
      }
    }
  }

  await pptx.writeFile({ fileName: "GiperLaser_Extra_Presentation.pptx" });
}

// Slide preview card component
const SlidePreview = ({ slide, index }: { slide: (typeof slides)[0]; index: number }) => {
  return (
    <div className="slide-card group relative bg-white rounded-sm overflow-hidden shadow-md hover:shadow-xl transition-all duration-300 hover:-translate-y-1 cursor-default">
      <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-red-700 z-10" />

      {slide.type === "cover" ? (
        <div className="aspect-[16/9] bg-gray-900 relative overflow-hidden">
          <img
            src={LASER_IMAGE}
            alt="laser"
            className="absolute right-0 top-0 w-3/5 h-full object-cover opacity-60"
          />
          <div className="absolute inset-0 bg-gradient-to-r from-gray-900 via-gray-900/90 to-transparent" />
          <div className="absolute left-5 top-1/2 -translate-y-1/2">
            <div className="text-white font-black font-montserrat text-xl leading-tight">
              GiperLaser<br />
              <span className="text-red-500">Extra</span>
            </div>
            <div className="w-10 h-0.5 bg-red-600 my-1.5" />
            <div className="text-gray-300 text-[9px] font-ibm leading-tight">
              Разработка инновационных<br />машин термической резки
            </div>
            <div className="text-red-400 text-[8px] font-montserrat font-bold mt-1.5 tracking-wide">
              НИОКР · 2025–2027
            </div>
          </div>
        </div>
      ) : (
        <div className="aspect-[16/9] bg-white relative overflow-hidden">
          {/* Red header */}
          <div className="absolute top-0 left-1.5 right-0 h-[22%] bg-red-700 flex items-center px-3">
            <span className="text-white font-black font-montserrat text-[11px] leading-tight truncate">
              {slide.title}
            </span>
          </div>

          {/* Content area */}
          <div className="absolute top-[24%] left-3 right-2 bottom-[12%] overflow-hidden">
            {slide.type === "points" && slide.points && (
              <div className="space-y-1">
                {slide.points.map((pt, i) => (
                  <div key={i} className="flex items-start gap-1.5">
                    <div className="w-1.5 h-1.5 bg-red-600 rounded-none mt-1 shrink-0" />
                    <span className="text-gray-800 text-[8px] font-ibm leading-tight">{pt}</span>
                  </div>
                ))}
              </div>
            )}
            {slide.type === "cols" && slide.cols && (
              <div className="flex gap-1.5 h-full">
                {slide.cols.map((col, ci) => (
                  <div key={ci} className="flex-1 bg-gray-50 rounded-sm overflow-hidden">
                    <div className="bg-red-600 px-1.5 py-0.5">
                      <span className="text-white text-[7px] font-montserrat font-bold">{col.label}</span>
                    </div>
                    <div className="p-1.5 space-y-0.5">
                      {col.items.map((item, ii) => (
                        <div key={ii} className="text-gray-700 text-[7px] font-ibm">• {item}</div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            )}
            {slide.type === "stats" && slide.stats && (
              <div className="grid grid-cols-2 gap-1.5 h-full">
                {slide.stats.map((st, si) => (
                  <div key={si} className="bg-gray-50 border-l-2 border-red-600 px-1.5 py-1 flex flex-col justify-center">
                    <div className="text-red-700 font-black font-montserrat text-[11px] leading-tight">{st.value}</div>
                    <div className="text-gray-600 text-[7px] font-ibm leading-tight mt-0.5">{st.label}</div>
                  </div>
                ))}
              </div>
            )}
            {slide.type === "final" && (
              <div className="space-y-1">
                {slide.subtitle && (
                  <div className="text-gray-500 text-[8px] font-ibm mb-1">{slide.subtitle}</div>
                )}
                {slide.points?.map((pt, i) => (
                  <div key={i} className="flex items-start gap-1.5">
                    <div className="w-1.5 h-1.5 bg-red-600 mt-1 shrink-0" />
                    <span className="text-gray-800 text-[8px] font-ibm leading-tight">{pt}</span>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Footer */}
          <div className="absolute bottom-0 left-1.5 right-0 h-[10%] bg-gray-100 flex items-center justify-between px-2">
            <span className="text-gray-400 text-[6px] font-ibm">ООО «Гиперплазма» · GiperLaser Extra</span>
            <span className="text-red-600 text-[7px] font-montserrat font-bold">{index + 1}/10</span>
          </div>
        </div>
      )}

      {/* Slide number badge */}
      <div className="absolute top-2 right-2 bg-red-700 text-white text-[9px] font-montserrat font-bold w-5 h-5 rounded-full flex items-center justify-center z-20">
        {index + 1}
      </div>
    </div>
  );
};

const Index = () => {
  const [generating, setGenerating] = useState(false);

  const handleDownload = async () => {
    setGenerating(true);
    try {
      await generatePptx();
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#F4F4F4] font-ibm">
      {/* Top bar */}
      <div className="bg-[#1A1A1A] border-b-4 border-red-700">
        <div className="max-w-7xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="w-1 h-8 bg-red-700 rounded-none" />
            <div>
              <div className="text-white font-black font-montserrat text-xl tracking-tight leading-none">
                GiperLaser <span className="text-red-500">Extra</span>
              </div>
              <div className="text-gray-400 text-xs font-ibm mt-0.5">
                Инновационные машины лазерной резки металлов
              </div>
            </div>
          </div>
          <button
            onClick={handleDownload}
            disabled={generating}
            className="flex items-center gap-2.5 bg-red-700 hover:bg-red-800 disabled:bg-red-900 text-white font-montserrat font-bold text-sm px-6 py-3 transition-all duration-200 hover:shadow-lg hover:shadow-red-900/40 active:scale-95 disabled:opacity-70"
          >
            <Icon name={generating ? "Loader2" : "Download"} size={16} className={generating ? "animate-spin" : ""} />
            {generating ? "Генерация..." : "Скачать PPTX"}
          </button>
        </div>
      </div>

      {/* Hero */}
      <div className="bg-white border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-6 py-8 flex items-center justify-between">
          <div>
            <h1 className="text-3xl font-black font-montserrat text-gray-900 leading-tight">
              Презентация проекта
            </h1>
            <p className="text-gray-500 font-ibm mt-1 text-sm">
              10 слайдов · PowerPoint · Строгий деловой стиль
            </p>
          </div>
          <div className="flex gap-6 text-center">
            {[
              { val: "10", label: "слайдов" },
              { val: "PPTX", label: "формат" },
              { val: "16:9", label: "пропорции" },
            ].map((item) => (
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
          {slides.map((slide, i) => (
            <SlidePreview key={slide.id} slide={slide} index={i} />
          ))}
        </div>
      </div>

      {/* Bottom CTA */}
      <div className="border-t border-gray-200 bg-white mt-4">
        <div className="max-w-7xl mx-auto px-6 py-8 flex flex-col sm:flex-row items-center justify-between gap-4">
          <div className="text-sm text-gray-500 font-ibm text-center sm:text-left">
            Файл формата <strong>.pptx</strong> — открывается в PowerPoint, Keynote, Google Slides
          </div>
          <button
            onClick={handleDownload}
            disabled={generating}
            className="flex items-center gap-2.5 bg-red-700 hover:bg-red-800 disabled:bg-red-900 text-white font-montserrat font-bold text-sm px-8 py-3.5 transition-all duration-200 hover:shadow-lg hover:shadow-red-900/40 active:scale-95 disabled:opacity-70"
          >
            <Icon name={generating ? "Loader2" : "FileDown"} size={16} className={generating ? "animate-spin" : ""} />
            {generating ? "Генерация файла..." : "Скачать презентацию"}
          </button>
        </div>
      </div>
    </div>
  );
};

export default Index;
