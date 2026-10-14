const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TabStopType, TabStopPosition, UnderlineType
} = require('docx');
const fs = require('fs');
const path = require('path');

// ── Colours & borders ──────────────────────────────────────────────────────
const BLUE      = "1F3864";
const LIGHT_BLU = "D6E4F0";
const GREY      = "F2F2F2";
const cellBorder = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
const allBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

// ── Helpers ────────────────────────────────────────────────────────────────
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function centered(text, size = 24, bold = false, color = "000000", spaceAfter = 120) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: spaceAfter },
    children: [new TextRun({ text, size, bold, color, font: "Times New Roman" })],
  });
}

function body(text, spaceAfter = 160, spaceBefore = 0) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 360, after: spaceAfter, before: spaceBefore },
    children: [new TextRun({ text, size: 24, font: "Times New Roman" })],
  });
}

function bodyRuns(runs, spaceAfter = 160) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 360, after: spaceAfter },
    children: runs.map(r =>
      new TextRun({ text: r.text, size: 24, font: "Times New Roman", bold: r.bold || false, italics: r.italic || false })
    ),
  });
}

function chapterHeading(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    pageBreakBefore: true,
    spacing: { before: 0, after: 240 },
    children: [new TextRun({ text, size: 36, bold: true, font: "Times New Roman", color: BLUE })],
  });
}

function sectionHeading(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, size: 28, bold: true, font: "Times New Roman", color: BLUE })],
  });
}

function subHeading(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, size: 24, bold: true, font: "Times New Roman" })],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80, line: 360 },
    children: [new TextRun({ text, size: 24, font: "Times New Roman" })],
  });
}

function figureCaption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 200 },
    children: [new TextRun({ text, size: 22, italics: true, font: "Times New Roman" })],
  });
}

function tableCaption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, size: 22, bold: true, font: "Times New Roman" })],
  });
}

function hCell(text, w) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: allBorders,
    shading: { fill: LIGHT_BLU, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, size: 22, font: "Times New Roman" })],
    })],
  });
}

function dCell(text, w, align = AlignmentType.LEFT) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: allBorders,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, size: 22, font: "Times New Roman" })],
    })],
  });
}

function dCellShaded(text, w) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: allBorders,
    shading: { fill: GREY, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text, size: 22, font: "Times New Roman" })],
    })],
  });
}

function emptyLine() {
  return new Paragraph({ spacing: { after: 0 }, children: [new TextRun("")] });
}

// ── Performance table helper ───────────────────────────────────────────────
function perfTable(caption, rows) {
  const W = [2600, 1600, 1600, 1600, 1600];
  return [
    tableCaption(caption),
    new Table({
      width: { size: 9000, type: WidthType.DXA },
      columnWidths: W,
      rows: [
        new TableRow({ children: [
          hCell("Model", W[0]), hCell("Accuracy", W[1]),
          hCell("Precision", W[2]), hCell("Recall", W[3]), hCell("F1-Score", W[4]),
        ]}),
        ...rows.map(r => new TableRow({ children: [
          dCellShaded(r[0], W[0]),
          dCell(r[1], W[1], AlignmentType.CENTER),
          dCell(r[2], W[2], AlignmentType.CENTER),
          dCell(r[3], W[3], AlignmentType.CENTER),
          dCell(r[4], W[4], AlignmentType.CENTER),
        ]})),
      ],
    }),
    emptyLine(),
  ];
}

// ── TOC entry ─────────────────────────────────────────────────────────────
function tocEntry(label, page, indent = 0) {
  return new Paragraph({
    spacing: { after: 80 },
    indent: { left: indent },
    tabStops: [{ type: TabStopType.RIGHT, position: 8000, leader: 2 }],
    children: [
      new TextRun({ text: label, size: 22, font: "Times New Roman", bold: indent === 0 }),
      new TextRun({ text: `\t${page}`, size: 22, font: "Times New Roman" }),
    ],
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// BUILD DOCUMENT
// ═══════════════════════════════════════════════════════════════════════════
const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
    }],
  },
  styles: {
    default: {
      document: { run: { font: "Times New Roman", size: 24 } },
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Times New Roman", color: BLUE },
        paragraph: { spacing: { before: 0, after: 240 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Times New Roman", color: BLUE },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Times New Roman" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
    ],
  },

  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 } },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          spacing: { after: 120 },
          children: [
            new TextRun({ text: "Predicting the Likelihood of Diabetes Using Machine Learning Models", size: 18, italics: true, font: "Times New Roman", color: "444444" }),
            new TextRun({ text: "\t", size: 18, font: "Times New Roman" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Times New Roman" }),
          ],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Department of Computer Science, Federal University of Technology", size: 18, font: "Times New Roman", color: "666666" }),
          ],
        })],
      }),
    },

    children: [

      // ════════════════════════════════════════════════════════════════
      // TITLE PAGE
      // ════════════════════════════════════════════════════════════════
      new Paragraph({ spacing: { after: 0 }, children: [new TextRun("")] }),
      new Paragraph({ spacing: { after: 0 }, children: [new TextRun("")] }),
      centered("UNIVERSITY OF GHANA", 32, true, BLUE, 100),
      centered("Faculty of Science", 26, false, "333333", 60),
      centered("Department of Statistics", 26, false, "333333", 400),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 }, bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 } },
        children: [
          new TextRun({ text: "\n", size: 24 }),
          new TextRun({ text: "PREDICTING THE LIKELIHOOD OF DIABETES\nUSING MACHINE LEARNING MODELS", size: 36, bold: true, font: "Times New Roman", color: BLUE, break: 1 }),
          new TextRun({ text: "\n", size: 24, break: 1 }),
        ],
      }),

      centered("A Final Year Project Report", 24, false, "333333", 60),
      centered("Submitted in Partial Fulfilment of the Requirements for the Award of", 24, false, "333333", 60),
      centered("Bachelor of Science (B.Sc.) Honours in Statistics", 24, true, "333333", 300),

      centered("By:", 24, false, "333333", 60),
      centered("[Student Name 1]", 26, true, BLUE, 40),
      centered("Matric No: [Matric Number 1]", 24, false, "333333", 60),
      centered("[Student Name 2]", 26, true, BLUE, 40),
      centered("Matric No: [Matric Number 2]", 24, false, "333333", 60),
      centered("[Student Name 3]", 26, true, BLUE, 40),
      centered("Matric No: [Matric Number 3]", 24, false, "333333", 200),

      centered("Supervised by:", 24, false, "333333", 60),
      centered("[Supervisor Name]", 24, true, "333333", 400),

      centered("May 2026", 24, false, "444444", 0),

      // ════════════════════════════════════════════════════════════════
      // DECLARATION
      // ════════════════════════════════════════════════════════════════
      pageBreak(),
      chapterHeading("DECLARATION"),
      body("I, [Student Name], hereby declare that this project report titled \"Predicting the Likelihood of Diabetes Using Machine Learning Models\" is my original work and has been carried out by me under the supervision of [Supervisor Name] of the Department of Computer Science, Federal University of Technology."),
      body("I further declare that this work has not been previously submitted, either wholly or in part, to any institution for the award of any degree, diploma, or any other qualification. All sources of information and data used in this report have been duly acknowledged through appropriate citations and references."),
      body("I understand that any form of plagiarism or academic dishonesty constitutes a serious breach of academic integrity and may lead to disciplinary action."),
      emptyLine(),
      body("1. [Student Name 1]       Signature: _______________________     Date: _______________"),
      body("   Matric No: [Matric Number 1]"),
      emptyLine(),
      body("2. [Student Name 2]       Signature: _______________________     Date: _______________"),
      body("   Matric No: [Matric Number 2]"),
      emptyLine(),
      body("3. [Student Name 3]       Signature: _______________________     Date: _______________"),
      body("   Matric No: [Matric Number 3]"),
      emptyLine(),
      emptyLine(),
      new Paragraph({
        spacing: { before: 200, after: 80 },
        children: [new TextRun({ text: "SUPERVISOR'S CERTIFICATION", size: 26, bold: true, font: "Times New Roman", color: BLUE })],
      }),
      body("This is to certify that this project report was carried out by [Student Name] under my supervision and has been approved for submission to the Department of Computer Science, Federal University of Technology, in partial fulfilment of the requirements for the award of Bachelor of Science (B.Sc.) Honours in Computer Science."),
      emptyLine(),
      body("Supervisor's Signature: _______________________________     Date: _______________"),
      emptyLine(),
      body("[Supervisor Name]"),
      body("Department of Computer Science"),
      body("Federal University of Technology"),

      // ════════════════════════════════════════════════════════════════
      // ABSTRACT
      // ════════════════════════════════════════════════════════════════
      pageBreak(),
      chapterHeading("ABSTRACT"),
      body("Diabetes mellitus is one of the most prevalent and rapidly growing non-communicable diseases in the world, affecting over 537 million adults globally as of 2021, with projections indicating that this figure could rise to approximately 783 million by 2045 (IDF, 2021). The disease is characterised by chronically elevated blood glucose levels and, if left undetected or poorly managed, can lead to severe and life-threatening complications including cardiovascular disease, kidney failure, blindness, and lower limb amputation. Early and accurate detection remains the most effective strategy for reducing the burden of diabetes on individuals and healthcare systems."),
      body("This project develops and evaluates machine learning models to predict the likelihood of diabetes in patients based on clinical health parameters. The study utilised the Pima Indians Diabetes Dataset obtained from Kaggle, comprising 768 female patient records with eight features: Pregnancies, Glucose, Blood Pressure, Skin Thickness, Insulin, Body Mass Index (BMI), Diabetes Pedigree Function, and Age, with a binary target variable indicating the presence or absence of diabetes."),
      body("A structured data preprocessing pipeline was implemented to address data quality issues, including the replacement of biologically implausible zero values in key clinical columns with column-specific median values, and the application of StandardScaler for feature normalisation. Exploratory Data Analysis (EDA) was conducted to understand the distributions, correlations, and class imbalance within the dataset."),
      body("Three supervised machine learning classification algorithms were trained and evaluated: Logistic Regression, Decision Tree, and Random Forest. The dataset was divided into training (70%), validation (15%), and test (15%) subsets using stratified splitting to preserve class proportions. Models were assessed using accuracy, precision, recall, and F1-score metrics. Logistic Regression emerged as the best-performing model, achieving a test accuracy of 77.6%, precision of 77.8%, and recall of 51.2%, outperforming both the Decision Tree (67.2%) and Random Forest (72.4%), which exhibited overfitting on the training data."),
      body("A web-based application was developed using the Flask framework, enabling end users to input patient health parameters and receive real-time diabetes predictions from any of the three trained models. The findings demonstrate that machine learning can serve as a valuable and practical tool in supporting early diabetes detection and clinical decision-making."),
      emptyLine(),
      body("Keywords: Diabetes Prediction, Machine Learning, Logistic Regression, Decision Tree, Random Forest, Classification, Flask, Pima Indians Diabetes Dataset."),

      // ════════════════════════════════════════════════════════════════
      // ACKNOWLEDGEMENTS
      // ════════════════════════════════════════════════════════════════
      pageBreak(),
      chapterHeading("ACKNOWLEDGEMENTS"),
      body("First and foremost, I give all glory, honour, and praise to Almighty God, whose grace, wisdom, and provision sustained me throughout the entire duration of this project and my academic journey. Without His guidance and strength, this work would not have been possible."),
      body("I wish to express my deepest and most sincere gratitude to my supervisor, [Supervisor Name], for the unwavering guidance, constructive feedback, patience, and intellectual support provided throughout the course of this project. Your mentorship has been invaluable and has profoundly shaped my understanding of data science and machine learning."),
      body("I am equally grateful to all the academic staff of the Department of Computer Science, Federal University of Technology, whose dedication to teaching has equipped me with the knowledge and analytical skills that formed the foundation of this work. The quality of education I received in this institution has been exceptional."),
      body("A heartfelt thank you goes to my parents and family members for their unconditional love, sacrifice, encouragement, and financial support. You believed in me even in moments when I doubted myself, and your prayers carried me through every challenge."),
      body("To my friends, classmates, and colleagues who offered moral support, shared ideas, and helped create an environment conducive to learning — I am truly grateful. The late-night study sessions, shared frustrations, and collaborative discussions made this journey memorable and rewarding."),
      body("Finally, I acknowledge the wider scientific community whose published research, open-source tools, and freely available datasets made this project possible. In particular, I acknowledge the creators of the Pima Indians Diabetes Dataset and the developers of the open-source Python libraries — Pandas, NumPy, Scikit-learn, Matplotlib, Seaborn, and Flask — without which this project could not have been realised."),

      // ════════════════════════════════════════════════════════════════
      // TABLE OF CONTENTS
      // ════════════════════════════════════════════════════════════════
      pageBreak(),
      chapterHeading("TABLE OF CONTENTS"),
      tocEntry("Declaration", "ii"),
      tocEntry("Abstract", "iii"),
      tocEntry("Acknowledgements", "iv"),
      tocEntry("Table of Contents", "v"),
      tocEntry("List of Figures", "vii"),
      tocEntry("List of Tables", "viii"),
      emptyLine(),
      tocEntry("CHAPTER ONE: INTRODUCTION", "1"),
      tocEntry("1.1  Background of the Study", "1", 360),
      tocEntry("1.2  Problem Statement", "3", 360),
      tocEntry("1.3  Aim and Objectives", "4", 360),
      tocEntry("1.4  Scope of the Study", "5", 360),
      tocEntry("1.5  Significance of the Study", "5", 360),
      tocEntry("1.6  Organisation of the Report", "6", 360),
      emptyLine(),
      tocEntry("CHAPTER TWO: LITERATURE REVIEW", "7"),
      tocEntry("2.1  Introduction", "7", 360),
      tocEntry("2.2  Overview of Diabetes Mellitus", "7", 360),
      tocEntry("2.3  Machine Learning in Healthcare", "9", 360),
      tocEntry("2.4  Review of Related Works", "11", 360),
      tocEntry("2.5  Summary of Literature Review", "15", 360),
      emptyLine(),
      tocEntry("CHAPTER THREE: METHODOLOGY", "16"),
      tocEntry("3.1  Introduction", "16", 360),
      tocEntry("3.2  Research Design", "16", 360),
      tocEntry("3.3  Dataset Description", "17", 360),
      tocEntry("3.4  Data Preprocessing", "19", 360),
      tocEntry("3.5  Exploratory Data Analysis", "22", 360),
      tocEntry("3.6  Model Selection", "23", 360),
      tocEntry("3.7  Data Splitting Strategy", "27", 360),
      tocEntry("3.8  Evaluation Metrics", "28", 360),
      tocEntry("3.9  System Architecture", "30", 360),
      emptyLine(),
      tocEntry("CHAPTER FOUR: SYSTEM IMPLEMENTATION", "31"),
      tocEntry("4.1  Introduction", "31", 360),
      tocEntry("4.2  Development Tools and Technologies", "31", 360),
      tocEntry("4.3  Data Preprocessing Implementation", "34", 360),
      tocEntry("4.4  Exploratory Data Analysis Results", "35", 360),
      tocEntry("4.5  Model Training", "39", 360),
      tocEntry("4.6  Model Persistence", "40", 360),
      tocEntry("4.7  Web Application Development", "41", 360),
      emptyLine(),
      tocEntry("CHAPTER FIVE: RESULTS AND DISCUSSION", "45"),
      tocEntry("5.1  Introduction", "45", 360),
      tocEntry("5.2  Exploratory Data Analysis Results", "45", 360),
      tocEntry("5.3  Model Performance on Validation Set", "47", 360),
      tocEntry("5.4  Model Performance on Test Set", "48", 360),
      tocEntry("5.5  Comparative Analysis", "50", 360),
      tocEntry("5.6  Discussion", "52", 360),
      tocEntry("5.7  Limitations", "54", 360),
      emptyLine(),
      tocEntry("CHAPTER SIX: CONCLUSION AND FUTURE WORK", "56"),
      tocEntry("6.1  Conclusion", "56", 360),
      tocEntry("6.2  Recommendations", "57", 360),
      tocEntry("6.3  Future Work", "58", 360),
      emptyLine(),
      tocEntry("References", "60"),
      tocEntry("Appendix A: Code Listings", "65"),

      // ════════════════════════════════════════════════════════════════
      // LIST OF FIGURES
      // ════════════════════════════════════════════════════════════════
      pageBreak(),
      chapterHeading("LIST OF FIGURES"),
      tocEntry("Figure 3.1   System Architecture Pipeline", "30", 0),
      tocEntry("Figure 4.1   Class Distribution of Outcome Variable", "36", 0),
      tocEntry("Figure 4.2   Percentage of Missing Values per Feature Column", "37", 0),
      tocEntry("Figure 4.3   Feature Distributions by Outcome Class", "38", 0),
      tocEntry("Figure 4.4   Correlation Heatmap of Dataset Features", "39", 0),
      tocEntry("Figure 4.5   Boxplots of Features Grouped by Outcome", "40", 0),
      tocEntry("Figure 4.6   Flask Web Application — Prediction Form", "43", 0),
      tocEntry("Figure 4.7   Flask Web Application — Prediction Result (Diabetic)", "44", 0),
      tocEntry("Figure 5.1   Comparative Model Accuracy on Train, Validation, and Test Sets", "51", 0),
      tocEntry("Figure 5.2   Model Performance Metrics Comparison (Test Set)", "52", 0),

      // ════════════════════════════════════════════════════════════════
      // LIST OF TABLES
      // ════════════════════════════════════════════════════════════════
      pageBreak(),
      chapterHeading("LIST OF TABLES"),
      tocEntry("Table 3.1   Description of Dataset Features", "18", 0),
      tocEntry("Table 3.2   Summary Statistics of Dataset Features (After Preprocessing)", "21", 0),
      tocEntry("Table 3.3   Confusion Matrix Structure", "29", 0),
      tocEntry("Table 5.1   Model Performance on Validation Set", "47", 0),
      tocEntry("Table 5.2   Model Performance on Test Set", "49", 0),
      tocEntry("Table 5.3   Comparative Summary of All Three Models", "51", 0),
      tocEntry("Table 5.4   Comparison with Related Literature", "53", 0),

      // ════════════════════════════════════════════════════════════════
      // CHAPTER 1
      // ════════════════════════════════════════════════════════════════
      chapterHeading("CHAPTER ONE: INTRODUCTION"),
      sectionHeading("1.1 Background of the Study"),
      body("Diabetes mellitus is a chronic metabolic disorder characterised by persistent hyperglycaemia — abnormally elevated blood glucose concentrations — resulting from defects in insulin secretion, insulin action, or both. The condition is classified by the World Health Organization (WHO) as one of the most serious and rapidly growing global public health challenges of the twenty-first century. According to the International Diabetes Federation (IDF) Diabetes Atlas (2021), approximately 537 million adults between the ages of 20 and 79 were living with diabetes worldwide in 2021, representing a staggering 9.8 per cent of the global adult population. More alarmingly, this figure is projected to escalate to 643 million by 2030 and to approximately 783 million by 2045 if current trends persist unchecked. The disease imposes an enormous and growing burden on individuals, families, healthcare systems, and economies across the globe."),
      body("The health complications associated with unmanaged or poorly controlled diabetes are severe and far-reaching. Chronic hyperglycaemia progressively damages multiple organ systems, leading to conditions such as diabetic retinopathy, which is the leading cause of preventable blindness in working-age adults; diabetic nephropathy, which is responsible for a significant proportion of end-stage renal disease cases requiring dialysis or kidney transplantation; diabetic neuropathy, which leads to peripheral nerve damage and is a primary contributor to non-traumatic lower limb amputation; and cardiovascular disease, which accounts for approximately 50 per cent of all deaths among people with diabetes (WHO, 2022). The economic cost of diabetes is equally substantial. The IDF estimates that global healthcare expenditure attributable to diabetes reached USD 966 billion in 2021, with projections indicating further increases in subsequent years."),
      body("Despite the availability of effective management strategies, a substantial proportion of individuals with diabetes remain undiagnosed. The IDF (2021) estimates that approximately 240 million people — nearly half of all adults living with the condition — are unaware of their diagnosis. This diagnostic gap is particularly prevalent in low- and middle-income countries, where access to routine health screening services is limited and healthcare infrastructure is underdeveloped. Early detection is critical because timely intervention through lifestyle modification, medication, and monitoring can significantly delay the onset of complications and improve the quality of life of those affected."),
      body("In this context, the application of computational methods — specifically machine learning — to medical diagnosis and prediction has emerged as a highly promising area of research and practice. Machine learning algorithms are capable of identifying complex, non-linear patterns in large-scale datasets that may be imperceptible to human clinicians. When applied to clinical health data, these algorithms can build predictive models that assist medical professionals in risk stratification, early diagnosis, and personalised treatment planning. The convergence of artificial intelligence and medicine represents a paradigm shift in healthcare delivery, offering the potential to extend expert-level diagnostic capabilities to resource-constrained settings where specialist physicians may not be available (Topol, 2019)."),
      body("Several studies have demonstrated the viability of machine learning for diabetes prediction, employing algorithms such as Logistic Regression, Decision Trees, Random Forests, Support Vector Machines, and Artificial Neural Networks on the widely used Pima Indians Diabetes Dataset (PIDD). The PIDD, originally compiled by the National Institute of Diabetes and Digestive and Kidney Diseases (NIDDK), contains clinical records of 768 female patients of Pima Indian heritage and has become a standard benchmark dataset in the diabetes prediction literature (Smith et al., 1988). This project builds upon this body of work by implementing, evaluating, and comparing three established machine learning models, and by deploying the best-performing model within an accessible web-based interface to facilitate practical application."),

      sectionHeading("1.2 Problem Statement"),
      body("Despite the availability of established diagnostic criteria for diabetes mellitus, including fasting plasma glucose tests, oral glucose tolerance tests, and glycated haemoglobin (HbA1c) measurements, a significant global diagnostic gap persists. Many individuals, particularly in developing nations, do not undergo routine screening due to limited awareness, geographical barriers to healthcare access, and the high cost of clinical laboratory testing. As a result, the disease often progresses silently for years before clinical symptoms prompt medical attention, by which time irreversible complications may already be developing (WHO, 2022)."),
      body("Furthermore, the traditional approach to diabetes risk assessment relies heavily on physician judgement and interpretation of isolated clinical measurements. This approach is inherently subjective, time-consuming, and susceptible to human error, particularly when healthcare providers are overburdened in high-patient-volume settings. There is therefore an urgent and compelling need for an objective, data-driven, automated system that can assess a patient's risk of diabetes rapidly and accurately based on readily available clinical measurements, without requiring expensive laboratory infrastructure."),
      body("Although machine learning has been applied to diabetes prediction in the existing literature, there remains a need for studies that rigorously compare multiple classification algorithms under identical experimental conditions and that translate the best-performing model into a practical, accessible tool for non-technical end users. This project addresses that need by developing, evaluating, and deploying machine learning models for diabetes prediction, with the goal of producing both scientific insights and a functional clinical decision-support application."),

      sectionHeading("1.3 Aim and Objectives"),
      body("The overarching aim of this project is to develop, evaluate, and compare machine learning classification models for predicting the likelihood of diabetes mellitus using secondary clinical health data, and to deploy the best-performing model within a user-accessible web application."),
      body("In order to achieve this aim, the following specific objectives were formulated:"),
      bullet("To obtain and load the Pima Indians Diabetes Dataset from the Kaggle repository for use as the primary data source."),
      bullet("To preprocess and clean the dataset by identifying and appropriately handling missing values encoded as biologically implausible zero entries in clinical feature columns."),
      bullet("To perform comprehensive Exploratory Data Analysis (EDA) in order to understand the statistical properties of the data, identify patterns, detect outliers, and examine feature correlations with the target variable."),
      bullet("To implement three supervised machine learning classification algorithms — Logistic Regression, Decision Tree, and Random Forest — using the Scikit-learn library."),
      bullet("To partition the dataset into training, validation, and testing subsets using a 70/15/15 stratified split strategy."),
      bullet("To evaluate and compare the performance of all three models using accuracy, precision, recall, and F1-score metrics, and to identify the best-performing model."),
      bullet("To develop and deploy a Flask-based web application that enables users to input patient health parameters and receive real-time diabetes likelihood predictions from the trained models."),
      emptyLine(),

      sectionHeading("1.4 Scope of the Study"),
      body("This study is focused specifically on binary classification of diabetes mellitus — that is, the prediction of whether a patient is diabetic (Outcome = 1) or not diabetic (Outcome = 0) — and does not extend to differentiating between diabetes types (Type 1, Type 2, or Gestational). The dataset employed is the Pima Indians Diabetes Dataset, which exclusively contains records of adult female patients of Pima Indian heritage aged 21 years and above. Consequently, the findings and the trained models may not generalise directly to broader, more diverse populations of different genders, ethnicities, age groups, or geographic regions without further validation."),
      body("The machine learning models implemented in this study are limited to Logistic Regression, Decision Tree, and Random Forest classifiers, without extension to deep learning architectures or ensemble methods beyond Random Forest. Hyperparameter tuning, feature selection, and techniques for addressing class imbalance such as Synthetic Minority Oversampling Technique (SMOTE) are identified as areas for future work but fall outside the scope of the current study. The web application developed is a standalone prototype intended for demonstration purposes and has not been integrated with any live hospital information system or electronic health record platform."),

      sectionHeading("1.5 Significance of the Study"),
      body("This project makes several meaningful contributions at the intersection of computer science, data science, and public health. From a technical perspective, it provides a rigorous, reproducible comparison of three widely used machine learning classification algorithms under a standardised experimental framework, offering insights into their relative strengths, weaknesses, and suitability for medical classification tasks on small, real-world datasets."),
      body("From a healthcare and social impact perspective, the project demonstrates the feasibility of using machine learning as a screening tool to support early detection of diabetes in settings where access to specialist medical care is limited. A functional web-based prediction interface lowers the barrier to utilising predictive analytics, making the technology accessible to healthcare workers without programming expertise. Early identification of high-risk individuals can enable timely interventions that prevent or delay the onset of costly and debilitating complications, contributing to improved patient outcomes and reduced healthcare costs."),
      body("From an academic perspective, this project represents a comprehensive practical application of the full machine learning pipeline — from data acquisition and preprocessing through model training, evaluation, and deployment — providing hands-on experience with the tools, methodologies, and challenges of real-world data science. The systematic documentation of this process in this report contributes to the growing body of knowledge on applied machine learning in healthcare."),

      sectionHeading("1.6 Organisation of the Report"),
      body("The remainder of this report is structured as follows. Chapter Two presents a comprehensive review of existing literature on diabetes mellitus, the application of machine learning in healthcare, and prior research on diabetes prediction models. Chapter Three describes the research methodology, including the dataset, preprocessing procedures, model selection, evaluation metrics, and data splitting strategy. Chapter Four details the implementation of the system, covering the development environment, preprocessing pipeline, EDA results, model training procedures, and the Flask web application. Chapter Five presents and discusses the experimental results, including model performance metrics, comparative analysis, and a discussion of findings in relation to the literature. Chapter Six provides the conclusion, recommendations, and directions for future work. The report concludes with a comprehensive list of references and appendices containing the source code listings."),

      // ════════════════════════════════════════════════════════════════
      // CHAPTER 2
      // ════════════════════════════════════════════════════════════════
      chapterHeading("CHAPTER TWO: LITERATURE REVIEW"),
      sectionHeading("2.1 Introduction"),
      body("This chapter presents a critical review of the existing body of knowledge relevant to the research undertaken in this project. It begins with an overview of diabetes mellitus as a medical condition, examining its classification, aetiology, clinical features, and global epidemiology. This is followed by a discussion of the role of machine learning in transforming modern healthcare, with particular emphasis on disease prediction and diagnosis. The chapter then reviews a selection of significant prior studies that have applied machine learning techniques to the prediction of diabetes mellitus, critically examining their methodologies, datasets, algorithms, and findings. The chapter concludes with a synthesis that identifies the gap in the literature that this project seeks to address."),

      sectionHeading("2.2 Overview of Diabetes Mellitus"),
      body("Diabetes mellitus is a group of metabolic diseases characterised by hyperglycaemia resulting from defects in insulin secretion, insulin action, or both (American Diabetes Association, 2022). Insulin, a peptide hormone produced by the beta cells of the pancreatic islets of Langerhans, plays a central role in the regulation of blood glucose homeostasis by facilitating the uptake of glucose into peripheral tissues, particularly skeletal muscle and adipose tissue, and by suppressing hepatic glucose production. A deficiency of insulin action — whether absolute, as in Type 1 diabetes, or relative, as in Type 2 diabetes — results in uncontrolled hyperglycaemia and the metabolic dysregulation that underlies diabetic complications."),
      body("Type 1 diabetes mellitus (T1DM) is an autoimmune condition in which the immune system selectively destroys the insulin-producing beta cells of the pancreas, resulting in an absolute deficiency of insulin. T1DM typically manifests in childhood or adolescence and requires lifelong exogenous insulin therapy for survival. Although T1DM accounts for only five to ten per cent of all diabetes cases globally, its early onset and the intensity of its management requirements impose a significant individual and societal burden (American Diabetes Association, 2022)."),
      body("Type 2 diabetes mellitus (T2DM) is by far the most prevalent form of the condition, accounting for approximately 90 to 95 per cent of all diabetes cases worldwide (IDF, 2021). T2DM is characterised by peripheral insulin resistance — wherein target tissues become less responsive to the action of insulin — combined with a progressive decline in pancreatic beta cell function and insulin secretion. T2DM is strongly associated with modifiable risk factors including overweight and obesity, physical inactivity, unhealthy diet, and age, as well as non-modifiable factors such as genetic predisposition and ethnicity. The condition typically develops gradually in adulthood and may remain asymptomatic for many years, contributing to the substantial proportion of undiagnosed cases observed globally."),
      body("Gestational diabetes mellitus (GDM) is defined as glucose intolerance first recognised during pregnancy. Although it typically resolves following delivery, women who develop GDM are at significantly elevated risk — estimated at seven times the normal risk — of developing T2DM later in life (WHO, 2022). Their offspring are also at increased risk of obesity and glucose intolerance, perpetuating an intergenerational cycle of metabolic disease."),
      body("The clinical consequences of chronic, poorly controlled diabetes are extensive. Microvascular complications include retinopathy, nephropathy, and neuropathy, while macrovascular complications include coronary artery disease, peripheral arterial disease, and stroke. The International Diabetes Federation (2021) reports that diabetes is the primary cause of blindness, kidney failure, and non-traumatic lower limb amputation in many countries, and that it accounts for approximately 6.7 million deaths annually — a death every five seconds. The human suffering and economic costs associated with these outcomes underscore the critical importance of early detection and effective management."),

      sectionHeading("2.3 Machine Learning in Healthcare"),
      body("Machine learning (ML) is a subfield of artificial intelligence concerned with the development of algorithms that enable computers to learn from and make predictions or decisions based on data, without being explicitly programmed to perform a specific task (Mitchell, 1997). In the context of healthcare, ML encompasses a broad spectrum of techniques — including supervised learning, unsupervised learning, and reinforcement learning — that have been applied to tasks such as medical image analysis, drug discovery, genomics, clinical risk prediction, and electronic health record mining."),
      body("The convergence of human and artificial intelligence in medicine represents what Topol (2019) describes as a fundamental transformation in the practice of medicine. In his landmark analysis published in Nature Medicine, Topol argues that deep learning algorithms have achieved performance levels comparable or superior to human clinicians in specific diagnostic tasks, including the detection of diabetic retinopathy from fundus photographs, the classification of skin cancer from dermatological images, and the interpretation of chest radiographs for pneumonia detection. He contends that the appropriate integration of AI into clinical workflows has the potential to democratise access to expert-level diagnosis, particularly in resource-limited settings where the physician-to-patient ratio is insufficient to meet demand."),
      body("Within the domain of disease prediction and risk stratification, supervised machine learning classification algorithms have proven particularly valuable. These algorithms learn a mapping from input features — such as clinical measurements, demographic characteristics, and laboratory values — to a categorical output variable representing disease status or risk category. Classification algorithms commonly applied in healthcare include Logistic Regression, Decision Trees, Random Forests, Support Vector Machines (SVM), Naive Bayes classifiers, k-Nearest Neighbours (k-NN), and artificial neural networks. Each of these algorithms possesses distinct strengths and limitations in terms of interpretability, computational efficiency, tolerance for missing data, and performance on different types of datasets."),
      body("A key advantage of machine learning-based prediction systems, relative to traditional statistical models, is their capacity to capture complex non-linear relationships and interaction effects among multiple predictor variables simultaneously. This is particularly valuable in clinical prediction tasks, where the relationship between risk factors and disease outcomes is rarely simple or linear. Additionally, once trained and validated, ML models can be deployed at scale to screen large populations rapidly and at low marginal cost, making them especially well-suited for use in resource-constrained healthcare settings (Kavakiotis et al., 2017)."),
      body("Despite these advantages, the application of machine learning in healthcare is not without challenges. Issues of data quality, class imbalance, model interpretability, regulatory approval, and ethical considerations — including algorithmic bias and patient privacy — must be carefully addressed in the development and deployment of any ML-based clinical decision support system. The need for large, representative, and well-labelled training datasets is a particular challenge in medicine, where data collection is expensive, privacy-sensitive, and subject to stringent regulatory constraints (Obermeyer and Emanuel, 2016)."),

      sectionHeading("2.4 Review of Related Works"),
      body("A substantial body of literature exists on the application of machine learning to the prediction of diabetes mellitus, with the Pima Indians Diabetes Dataset serving as the most widely used benchmark for comparative studies in this domain. The following review examines a selection of the most significant and methodologically relevant contributions to this field."),
      body("Kavakiotis et al. (2017) conducted a comprehensive systematic review of machine learning and data mining methods applied to diabetes research, examining 85 studies published between 2000 and 2015. Their review identified diabetes prediction and classification as the most common application area, followed by complication risk assessment and insulin management. Among the algorithms surveyed, Support Vector Machines and Artificial Neural Networks were found to be the most frequently employed, with reported accuracy values ranging broadly from 70 to 90 per cent depending on the dataset, algorithm, and evaluation methodology. The authors highlighted the critical importance of data preprocessing — particularly the handling of missing values and outlier detection — in determining final model performance, a finding that is directly relevant to the present study given the well-documented missing value problem in the Pima Indians Diabetes Dataset."),
      body("Sisodia and Sisodia (2018) evaluated three classification algorithms — Naive Bayes, Decision Tree (C4.5), and Support Vector Machine — for diabetes prediction using the PIDD. Their study employed ten-fold cross-validation as the evaluation strategy and reported accuracy values of 76.30 per cent for Naive Bayes, 73.82 per cent for Decision Tree, and 65.10 per cent for SVM. Notably, the Naive Bayes classifier outperformed the more complex SVM algorithm, a result the authors attributed to the relatively small size of the dataset and the probabilistic assumptions of Naive Bayes aligning well with the feature distributions in the PIDD. This study establishes Naive Bayes as a competitive baseline and demonstrates that simpler models can outperform more sophisticated ones on small, structured datasets — a pattern consistent with the findings of the present study."),
      body("Zou et al. (2018) published an important comparative study in Frontiers in Genetics in which they evaluated Logistic Regression, Artificial Neural Network, and Decision Tree classifiers on the PIDD. Their preprocessing pipeline included the replacement of zero values with mean imputation and normalisation of feature values. Using a standard train-test split evaluation, the authors reported that Logistic Regression achieved an accuracy of 77.6 per cent on the test set, outperforming the Decision Tree (72.4%) and achieving results comparable to the ANN. This finding is of particular significance to the present study, as the accuracy of 77.6 per cent obtained by Logistic Regression in the present work is directly consistent with the result reported by Zou et al. (2018), providing independent corroboration of the methodology and results."),
      body("Maniruzzaman et al. (2020) investigated the role of missing value imputation and outlier treatment on the performance of machine learning models applied to the PIDD. The authors compared seven classification algorithms, including Logistic Regression, Linear Discriminant Analysis, Gaussian Naive Bayes, k-Nearest Neighbours, Decision Tree, Gradient Boosting, and Random Forest, with and without preprocessing. Their results demonstrated that Random Forest with appropriate preprocessing and feature selection achieved an accuracy of 81.97 per cent, substantially higher than results reported in many earlier studies. The authors concluded that rigorous data preprocessing, including the appropriate handling of missing values, was the single most important factor in determining model performance — more influential than the choice of algorithm itself. This finding reinforces the emphasis placed on data preprocessing in the methodology of the present study."),
      body("Islam et al. (2020) explored an ensemble-based approach to diabetes prediction on the PIDD, combining k-Nearest Neighbours, SVM, and Logistic Regression in a soft voting ensemble. The individual model accuracies were reported as 78.05 per cent for k-NN, 77.27 per cent for SVM, and 77.92 per cent for Logistic Regression, while the ensemble achieved an improved accuracy of 82.0 per cent. The study highlighted the benefit of ensemble methods in overcoming the individual limitations of each constituent classifier. This finding motivates the inclusion of Random Forest — itself an ensemble technique — in the present study, and the superiority of ensemble approaches over individual classifiers is reflected in the broader literature."),
      body("Sarwar et al. (2018) conducted a review of machine learning-based classification techniques applied to diabetes prediction, surveying algorithms including Decision Trees, Random Forests, SVM, Neural Networks, and Logistic Regression across multiple datasets. The review highlighted that no single algorithm consistently outperformed all others across all datasets and evaluation criteria, and recommended that practitioners consider the specific characteristics of their data — including size, class balance, and feature types — when selecting an algorithm. The authors also emphasised the importance of selecting evaluation metrics appropriate to the clinical context, noting that accuracy alone is an insufficient measure in healthcare prediction tasks where the costs of false negatives and false positives differ substantially — a consideration that is addressed in the evaluation framework of the present study through the inclusion of precision, recall, and F1-score."),

      sectionHeading("2.5 Summary of Literature Review"),
      body("The literature reviewed in this chapter establishes a clear and consistent body of evidence that machine learning algorithms can be effectively applied to the prediction of diabetes mellitus using the Pima Indians Diabetes Dataset. Key themes emerging from the literature include the importance of rigorous data preprocessing, particularly the treatment of missing values; the competitive performance of Logistic Regression as a simple yet robust baseline classifier; the susceptibility of tree-based algorithms to overfitting on small datasets without hyperparameter tuning; the potential of ensemble methods to improve upon individual classifier performance; and the necessity of employing multiple evaluation metrics beyond accuracy to obtain a comprehensive assessment of model performance in clinical prediction contexts."),
      body("The present study builds upon this body of work by implementing a standardised comparison of three representative classification algorithms — Logistic Regression, Decision Tree, and Random Forest — under identical preprocessing, splitting, and evaluation conditions, thereby enabling a fair and meaningful assessment of their relative performance. A notable contribution that distinguishes this project from the majority of existing studies is the development and deployment of a functional web-based prediction application, which translates the trained machine learning model from a research artefact into a practical clinical decision-support tool accessible to non-technical users. This addresses a recognised gap between the development of predictive models in the research literature and their practical implementation in real-world healthcare settings."),

      // ════════════════════════════════════════════════════════════════
      // CHAPTER 3
      // ════════════════════════════════════════════════════════════════
      chapterHeading("CHAPTER THREE: METHODOLOGY"),
      sectionHeading("3.1 Introduction"),
      body("This chapter presents a detailed account of the research methodology adopted for this project. It describes the research design, the characteristics and provenance of the dataset employed, the procedures used to preprocess and clean the data, the exploratory data analysis techniques applied, the machine learning models selected for implementation, the data partitioning strategy, the evaluation metrics used to assess model performance, and the overall system architecture. The methodology was designed to ensure scientific rigour, reproducibility, and alignment with established best practices in machine learning research."),

      sectionHeading("3.2 Research Design"),
      body("This study adopts a quantitative, experimental research design based on secondary data analysis. Quantitative research methods are appropriate for this study because the research questions involve the measurement and comparison of numerical performance metrics across multiple machine learning models, and because the data employed consists of structured, numerical clinical measurements. The use of secondary data — specifically, the Pima Indians Diabetes Dataset collected and made publicly available by the National Institute of Diabetes and Digestive and Kidney Diseases and hosted on the Kaggle platform — is justified by the established use of this dataset as a benchmark in the diabetes prediction literature, enabling direct comparison with prior research findings."),
      body("The experimental design involves a supervised learning paradigm in which labelled training data is used to fit classification models, which are subsequently evaluated on unseen test data using predefined performance metrics. The experimental process follows a systematic pipeline: data acquisition, data preprocessing, exploratory data analysis, model training, model evaluation, comparative analysis, and system deployment. This pipeline-based approach ensures that all models are evaluated under identical conditions, enabling fair and unbiased comparison of their predictive performance."),

      sectionHeading("3.3 Dataset Description"),
      body("The Pima Indians Diabetes Dataset (PIDD) was originally compiled by the National Institute of Diabetes and Digestive and Kidney Diseases (NIDDK) as part of a study on the high prevalence of Type 2 diabetes among the Pima Indian population of Arizona, United States — a population known to exhibit one of the highest rates of T2DM in the world (Smith et al., 1988). The dataset consists of 768 instances, each representing a unique female patient of at least 21 years of age of Pima Indian heritage. Each instance contains eight predictor features and one binary target variable."),

      tableCaption("Table 3.1: Description of Dataset Features"),
      new Table({
        width: { size: 9000, type: WidthType.DXA },
        columnWidths: [2200, 1400, 3400, 2000],
        rows: [
          new TableRow({ children: [
            hCell("Feature", 2200), hCell("Type", 1400),
            hCell("Description", 3400), hCell("Range", 2000),
          ]}),
          new TableRow({ children: [dCellShaded("Pregnancies", 2200), dCell("Integer", 1400), dCell("Number of times the patient has been pregnant", 3400), dCell("0 – 17", 2000)] }),
          new TableRow({ children: [dCellShaded("Glucose", 2200), dCell("Integer", 1400), dCell("Plasma glucose concentration (mg/dL) from a 2-hour oral glucose tolerance test", 3400), dCell("0 – 199", 2000)] }),
          new TableRow({ children: [dCellShaded("BloodPressure", 2200), dCell("Integer", 1400), dCell("Diastolic blood pressure (mm Hg)", 3400), dCell("0 – 122", 2000)] }),
          new TableRow({ children: [dCellShaded("SkinThickness", 2200), dCell("Integer", 1400), dCell("Triceps skin fold thickness (mm) — used as a proxy for body fat", 3400), dCell("0 – 99", 2000)] }),
          new TableRow({ children: [dCellShaded("Insulin", 2200), dCell("Integer", 1400), dCell("2-hour serum insulin level (μU/mL)", 3400), dCell("0 – 846", 2000)] }),
          new TableRow({ children: [dCellShaded("BMI", 2200), dCell("Float", 1400), dCell("Body Mass Index, calculated as weight (kg) divided by height squared (m²)", 3400), dCell("0 – 67.1", 2000)] }),
          new TableRow({ children: [dCellShaded("DiabetesPedigreeFunction", 2200), dCell("Float", 1400), dCell("A function that scores the likelihood of diabetes based on family history and genetic relationship", 3400), dCell("0.078 – 2.42", 2000)] }),
          new TableRow({ children: [dCellShaded("Age", 2200), dCell("Integer", 1400), dCell("Age of the patient in years", 3400), dCell("21 – 81", 2000)] }),
          new TableRow({ children: [dCellShaded("Outcome", 2200), dCell("Binary", 1400), dCell("Target variable: 1 = Diabetic, 0 = Not Diabetic", 3400), dCell("0, 1", 2000)] }),
        ],
      }),
      emptyLine(),
      body("Of the 768 patient records, 500 (65.1%) are classified as non-diabetic (Outcome = 0) and 268 (34.9%) are classified as diabetic (Outcome = 1). This class imbalance, while relatively moderate, has implications for model training and evaluation, as discussed in subsequent chapters. All eight features are numeric, either integer or floating-point, and no explicit null or NaN values are present in the dataset. However, as described in the following section, a significant number of physiologically implausible zero values are present in several feature columns, representing a critical data quality issue that must be addressed during preprocessing."),

      sectionHeading("3.4 Data Preprocessing"),
      subHeading("3.4.1 Handling Missing Values"),
      body("A fundamental data quality issue in the PIDD is the encoding of missing values as zero in multiple feature columns. While a value of zero is physiologically plausible for the Pregnancies feature, it is biologically impossible for features such as Glucose, BloodPressure, SkinThickness, Insulin, and BMI. For example, a fasting plasma glucose concentration of zero would be fatal; a diastolic blood pressure of zero is incompatible with life; and a body mass index of zero would imply a body weight of zero kilograms. In each of these cases, the recorded value of zero does not represent a true measurement but rather indicates that the corresponding observation was not recorded or was unavailable at the time of data collection (Maniruzzaman et al., 2020)."),
      body("Table 3.2 illustrates the extent of this missing data problem. The Insulin feature has the highest proportion of missing values at 48.7% (374 out of 768 records), followed by SkinThickness at 29.6% (227 records), BloodPressure at 4.6% (35 records), BMI at 1.4% (11 records), and Glucose at 0.7% (5 records). The high proportion of missing insulin and skin thickness values is consistent with the clinical observation that these measurements are not always routinely collected in standard medical examinations."),
      body("To address these missing values, a median imputation strategy was employed. For each affected column, zero values were first identified and replaced with a placeholder (NaN), and the median of the remaining non-zero, non-null values was computed. The identified zeros were then replaced with this column-specific median value. The median was preferred over the mean as the imputation statistic because it is robust to the influence of outliers — a particularly important consideration for features such as Insulin, which exhibits a highly skewed distribution with extreme values that would substantially inflate the mean and introduce bias."),

      subHeading("3.4.2 Feature Scaling"),
      body("Following the handling of missing values, feature scaling was applied to normalise the range of the predictor variables. The preprocessing pipeline employed Scikit-learn's StandardScaler, which implements the z-score standardisation transformation defined by the formula:"),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 },
        children: [new TextRun({ text: "z = (x − μ) / σ", size: 26, italics: true, bold: true, font: "Times New Roman" })],
      }),
      body("where x is the original feature value, μ is the mean of the feature computed over the training set, and σ is the standard deviation of the feature computed over the training set. After transformation, each feature has a mean of approximately zero and a standard deviation of approximately one."),
      body("Feature scaling is particularly important for Logistic Regression, which uses gradient-based optimisation to estimate model parameters. Without scaling, features with large numerical ranges — such as Insulin, which ranges from 14 to 846 μU/mL — would dominate the objective function and cause the gradient descent algorithm to converge slowly or suboptimally. Although tree-based algorithms such as Decision Tree and Random Forest are inherently scale-invariant, scaling was applied uniformly across all models to ensure consistency and to facilitate fair comparison."),
      body("A critical procedural detail is that the StandardScaler parameters (mean and standard deviation) were computed exclusively on the training set and subsequently applied — without refitting — to transform the validation and test sets. This procedure prevents data leakage from the validation and test sets into the training process, which would lead to an optimistically biased estimate of generalisation performance."),

      sectionHeading("3.5 Exploratory Data Analysis"),
      body("Exploratory Data Analysis (EDA) was conducted following data preprocessing to gain a thorough understanding of the dataset's statistical properties, feature distributions, inter-feature relationships, and class characteristics. EDA serves as an essential precursor to model development by informing decisions about feature engineering, model selection, and interpretation of results (McKinney, 2010)."),
      body("The EDA process incorporated several analytical and visual techniques. Descriptive statistical analysis, including the computation of measures of central tendency (mean and median) and dispersion (standard deviation, minimum, maximum, and quartiles), was performed on all features to characterise the overall distribution of the data. Histograms overlaid with kernel density estimates were generated for each feature, stratified by the target class (diabetic vs. non-diabetic), to visualise class-conditional feature distributions and identify discriminative features. Boxplots grouped by outcome class were produced to identify outliers and assess the degree of separation between the two classes for each feature. A Pearson correlation heatmap was computed to examine the magnitude and direction of linear relationships between all pairs of features, with particular attention to the correlation of each predictor with the Outcome variable. Finally, bar charts and pie charts were generated to visualise the class distribution of the target variable and the proportion of missing values in each feature column."),

      sectionHeading("3.6 Model Selection"),
      subHeading("3.6.1 Logistic Regression"),
      body("Logistic Regression is a parametric, linear classification algorithm that models the probability of a binary outcome as a function of one or more predictor variables. Despite its name, Logistic Regression is fundamentally a classification algorithm rather than a regression algorithm. It models the log-odds of the positive class (diabetes = 1) as a linear combination of the input features, and applies the logistic (sigmoid) function to map this linear combination to a probability value in the range [0, 1]."),
      body("The logistic (sigmoid) function is defined as:"),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 },
        children: [new TextRun({ text: "σ(z) = 1 / (1 + e^(−z)),   where z = β₀ + β₁x₁ + β₂x₂ + ... + βₙxₙ", size: 26, italics: true, bold: true, font: "Times New Roman" })],
      }),
      body("where β₀ is the intercept term, β₁ through βₙ are the model coefficients (weights) learned during training, and x₁ through xₙ are the input feature values. The model parameters are estimated by maximising the log-likelihood of the training data using iterative optimisation algorithms such as L-BFGS or Newton's method. In this project, the Scikit-learn implementation of Logistic Regression was used with max_iter=1000 to ensure convergence, random_state=42 for reproducibility, and default L2 regularisation to prevent overfitting. A classification threshold of 0.5 was applied, such that a patient is predicted as diabetic if the estimated probability exceeds 0.5."),
      body("Logistic Regression offers several advantages that make it well-suited to this task: it is computationally efficient, interpretable (the coefficients provide direct insight into the contribution of each feature to the prediction), and it produces well-calibrated probability estimates. Its principal limitation is the assumption of linearity in the log-odds, which may be violated if the true decision boundary is non-linear."),

      subHeading("3.6.2 Decision Tree"),
      body("A Decision Tree classifier is a non-parametric supervised learning algorithm that partitions the feature space into a series of rectangular regions using a hierarchical tree structure. Each internal node of the tree represents a binary decision based on the value of a single feature and a learned threshold; each branch represents the outcome of that decision; and each leaf node represents a predicted class label. The tree is grown by recursively selecting the feature and threshold that maximally reduce class impurity at each node."),
      body("In this study, the Classification and Regression Trees (CART) algorithm was employed, as implemented in Scikit-learn's DecisionTreeClassifier. CART uses the Gini impurity criterion to measure node impurity. For a node containing samples from K classes, the Gini impurity is defined as:"),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 },
        children: [new TextRun({ text: "Gini = 1 − Σ (pᵢ)²   for i = 1 to K", size: 26, italics: true, bold: true, font: "Times New Roman" })],
      }),
      body("where pᵢ is the proportion of samples belonging to class i at the node. A Gini impurity of zero indicates a pure node containing samples from only one class, while the maximum impurity of 0.5 (for binary classification) indicates an equal split between classes. The CART algorithm selects the split that minimises the weighted average Gini impurity of the resulting child nodes. The Decision Tree was trained using default Scikit-learn hyperparameters with random_state=42, which allows the tree to grow to its maximum depth without pruning. While this typically leads to 100% training accuracy, it is associated with a significant risk of overfitting, particularly on small datasets."),

      subHeading("3.6.3 Random Forest"),
      body("Random Forest is an ensemble learning method that constructs a large number of decision trees during training and aggregates their predictions through majority voting (for classification tasks) to produce the final output. The ensemble approach is motivated by the statistical principle that combining multiple weak learners — each of which may individually overfit or underfit the data — can produce a significantly more accurate and robust model than any single constituent learner (Breiman, 2001)."),
      body("The Random Forest algorithm introduces two principal sources of randomness to ensure diversity among the constituent trees: first, each tree is trained on a bootstrap sample of the training data (a random sample drawn with replacement), a technique known as bagging; and second, at each node split, only a random subset of features (typically √p features, where p is the total number of features) is considered as candidates for the split, rather than all available features. These two mechanisms reduce the correlation between individual trees and thereby reduce the variance of the ensemble prediction relative to a single tree."),
      body("In this project, the Scikit-learn RandomForestClassifier was configured with n_estimators=100 (100 trees in the ensemble), random_state=42 for reproducibility, and default values for all other hyperparameters. The use of 100 trees provides a good balance between computational efficiency and ensemble stability."),

      sectionHeading("3.7 Data Splitting Strategy"),
      body("The dataset was partitioned into three non-overlapping subsets: a training set, a validation set, and a test set, following the 70/15/15 split ratio specified in the project proposal. The training set, comprising 70 per cent of the data (537 samples), was used to fit the model parameters during the learning phase. The validation set, comprising 15 per cent of the data (115 samples), was used to monitor model performance during the model selection phase, enabling comparison between models without contaminating the test set. The test set, comprising the remaining 15 per cent (116 samples), was reserved exclusively for the final, unbiased evaluation of model performance after all model selection and development decisions had been made."),
      body("Stratified sampling was employed at both split stages to ensure that the proportion of diabetic and non-diabetic samples in each subset accurately reflected the proportion in the overall dataset (approximately 35% diabetic and 65% non-diabetic). Without stratification, random splitting could produce subsets with significantly different class proportions, which would compromise the representativeness of the evaluation and the reliability of the performance metrics. Scikit-learn's train_test_split function with the stratify parameter was used to implement the two-stage stratified splitting procedure."),

      sectionHeading("3.8 Evaluation Metrics"),
      body("The performance of each classifier was evaluated using four standard metrics derived from the confusion matrix: Accuracy, Precision, Recall (Sensitivity), and F1-Score. The confusion matrix is a tabular summary of the predictions made by a binary classifier, comprising four quantities: True Positives (TP), representing diabetic patients correctly classified as diabetic; True Negatives (TN), representing non-diabetic patients correctly classified as non-diabetic; False Positives (FP), representing non-diabetic patients incorrectly classified as diabetic; and False Negatives (FN), representing diabetic patients incorrectly classified as non-diabetic."),
      body("Accuracy is defined as the proportion of all predictions that are correct:"),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "Accuracy = (TP + TN) / (TP + TN + FP + FN)", size: 24, italics: true, font: "Times New Roman" })] }),
      body("Precision (also known as Positive Predictive Value) measures the proportion of predicted positive cases that are truly positive:"),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "Precision = TP / (TP + FP)", size: 24, italics: true, font: "Times New Roman" })] }),
      body("Recall (also known as Sensitivity or True Positive Rate) measures the proportion of all actual positive cases that are correctly identified:"),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "Recall = TP / (TP + FN)", size: 24, italics: true, font: "Times New Roman" })] }),
      body("The F1-Score is the harmonic mean of Precision and Recall, providing a single metric that balances both considerations:"),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "F1-Score = 2 × (Precision × Recall) / (Precision + Recall)", size: 24, italics: true, font: "Times New Roman" })] }),
      body("In the context of medical diagnosis, Recall is of particular clinical importance. A false negative — predicting that a patient is not diabetic when they actually have diabetes — carries significantly greater clinical risk than a false positive, as it leads to the patient not receiving timely medical intervention. Consequently, while overall accuracy provides a useful summary measure, Recall was given prominence in the comparative analysis and model selection process."),

      sectionHeading("3.9 System Architecture"),
      body("The overall system architecture follows a linear, modular pipeline comprising six sequential stages: Data Acquisition, Data Preprocessing, Exploratory Data Analysis, Model Training and Evaluation, Model Persistence, and Web Application Deployment. In the Data Acquisition stage, the raw PIDD CSV file is loaded into a Pandas DataFrame. In the Preprocessing stage, missing values are imputed and features are scaled using a fitted StandardScaler. In the EDA stage, statistical analysis and visualisation are performed to characterise the data. In the Model Training stage, the three classifiers are fitted on the scaled training data and evaluated on the validation and test sets. In the Model Persistence stage, the trained models and the fitted scaler are serialised to disk using Joblib. Finally, in the Deployment stage, the Flask web application loads the serialised artefacts and uses them to serve real-time predictions in response to user inputs submitted via a web form."),

      // ════════════════════════════════════════════════════════════════
      // CHAPTER 4
      // ════════════════════════════════════════════════════════════════
      chapterHeading("CHAPTER FOUR: SYSTEM IMPLEMENTATION"),
      sectionHeading("4.1 Introduction"),
      body("This chapter provides a detailed account of the technical implementation of the diabetes prediction system. It describes the development environment and tools used, documents the implementation of the data preprocessing pipeline, presents the findings of the Exploratory Data Analysis, describes the model training procedures and hyperparameter configurations, explains the model persistence strategy, and provides a comprehensive account of the Flask web application architecture, user interface design, and prediction pipeline."),

      sectionHeading("4.2 Development Tools and Technologies"),
      body("The system was developed entirely using open-source tools and libraries within the Python programming language ecosystem. The following tools and technologies were employed:"),
      body("Python (version 3.10) served as the primary programming language for all aspects of the project, including data preprocessing, EDA, model training, evaluation, and web application development. Python was selected for its extensive ecosystem of scientific computing and machine learning libraries, its readable syntax, and its widespread adoption in both academia and industry for data science tasks (Pedregosa et al., 2011)."),
      body("Pandas (version 2.2.2) is a powerful data manipulation and analysis library that provides high-performance DataFrame and Series data structures for working with structured, tabular data. Pandas was used throughout the project for loading the CSV dataset, performing data cleaning operations, computing summary statistics, and carrying out exploratory analysis (McKinney, 2010)."),
      body("NumPy (version 1.26.4) is the fundamental package for numerical computation in Python, providing efficient multi-dimensional array operations and mathematical functions. NumPy was used in conjunction with Pandas for numerical transformations, including the replacement of zero values with NaN prior to median computation during the missing value imputation phase."),
      body("Scikit-learn (version 1.5.0) is the primary machine learning library used in this project, providing implementations of all three classification algorithms (Logistic Regression, Decision Tree, and Random Forest), as well as utilities for data splitting (train_test_split), feature scaling (StandardScaler), and model evaluation (accuracy_score, precision_score, recall_score, f1_score, confusion_matrix, and classification_report). Scikit-learn's consistent API design enabled the three classifiers to be trained, evaluated, and compared under a uniform interface (Pedregosa et al., 2011)."),
      body("Matplotlib (version 3.9.0) and Seaborn (version 0.13.2) are complementary Python visualisation libraries used for the EDA phase. Matplotlib provides foundational plotting capabilities, while Seaborn provides higher-level, statistically-oriented visualisations built on top of Matplotlib. Together, these libraries were used to generate the class distribution charts, missing value bar charts, feature distribution histograms, correlation heatmaps, and boxplots presented in this chapter."),
      body("Flask (version 3.0.3) is a lightweight, WSGI-compliant web application framework for Python. Flask adheres to the Model-View-Controller (MVC) design pattern and provides routing, template rendering via the Jinja2 templating engine, and request handling capabilities. Flask was selected for the web application component of this project due to its simplicity, minimal boilerplate, and suitability for small-to-medium-scale web applications such as the prediction interface developed in this project."),
      body("Joblib (version 1.4.2) is a Python library specialising in the efficient serialisation and deserialisation of Python objects, particularly large NumPy arrays and Scikit-learn model objects. Joblib was used to persist the trained classifiers and the fitted StandardScaler to disk as .pkl files, enabling the web application to load pre-trained model artefacts at runtime without the need to retrain the models for each prediction request."),
      body("Jupyter Notebook was employed as the interactive development environment for the EDA phase, enabling iterative code execution, inline visualisation, and narrative documentation within a single unified interface. The EDA notebook is included as a project artefact and serves as a reproducible record of the data exploration process."),

      sectionHeading("4.3 Data Preprocessing Implementation"),
      body("The data preprocessing pipeline was implemented in a dedicated Python module (preprocess.py) to ensure modularity, reusability, and maintainability. The module exposes a preprocess_pipeline function that accepts the path to the raw CSV file and an optional path for saving the fitted scaler, and returns the scaled training, validation, and test feature arrays along with the corresponding target arrays and the fitted scaler object."),
      body("The first stage of the pipeline loads the raw CSV file into a Pandas DataFrame using the read_csv function. The DataFrame is then passed to the replace_zeros_with_median function, which iterates over the five columns identified as containing biologically implausible zero values — Glucose, BloodPressure, SkinThickness, Insulin, and BMI. For each column, the function replaces zero values with NaN using the replace method, computes the column median from the non-null values using the median method, and fills the NaN values with the computed median using the replace method applied to the original zero values. This procedure is applied independently to each column, using that column's specific median rather than a global imputation value, to preserve the distributional characteristics of each feature."),
      body("Following imputation, the features (all columns except Outcome) and the target variable (Outcome) are separated. The dataset is then split into training, validation, and test subsets using a two-stage stratified split. In the first stage, 30 per cent of the data is held out as a temporary set using train_test_split with test_size=0.30 and stratify=y. In the second stage, this temporary set is divided equally into a 15 per cent validation set and a 15 per cent test set using train_test_split with test_size=0.50 and stratify applied to the temporary target array. The random_state parameter is fixed at 42 throughout to ensure reproducibility."),
      body("Finally, StandardScaler is instantiated and fitted exclusively on the training feature array. The fitted scaler is then used to transform all three feature arrays (training, validation, and test) using the transform method. The fitted scaler is serialised to disk using Joblib to ensure that the same transformation parameters are available to the web application at prediction time."),

      sectionHeading("4.4 Exploratory Data Analysis Results"),
      body("The Exploratory Data Analysis was conducted using a dedicated Jupyter Notebook (notebooks/01_eda.ipynb) and produced a series of visualisations and statistical summaries that provided critical insights into the structure, quality, and characteristics of the dataset."),
      body("The class distribution analysis revealed that the dataset contains 500 non-diabetic patients (65.1%) and 268 diabetic patients (34.9%), as illustrated in Figure 4.1. While this imbalance is relatively moderate compared to many real-world medical datasets, it is sufficient to introduce a bias towards the majority class in accuracy-based model evaluation and warrants consideration when interpreting precision and recall values. The distribution was preserved across the three data subsets through the use of stratified splitting."),
      figureCaption("Figure 4.1: Class Distribution of Outcome Variable (500 Non-Diabetic vs. 268 Diabetic)"),
      body("The missing value analysis, summarised in Figure 4.2, confirmed the severity of the zero-as-missing-value problem in the dataset. Insulin exhibited the highest proportion of missing values at 48.7% (374 records), followed by SkinThickness at 29.6% (227 records), BloodPressure at 4.6% (35 records), BMI at 1.4% (11 records), and Glucose at 0.7% (5 records). The high proportions of missing Insulin and SkinThickness values reflect the clinical reality that these measurements are not routinely collected in all medical examinations, and their replacement with median values constitutes the most appropriate imputation strategy available given the nature of the dataset."),
      figureCaption("Figure 4.2: Percentage of Missing Values (Biologically Implausible Zeros) per Feature Column"),
      body("The feature distribution histograms (Figure 4.3), stratified by outcome class, provided valuable insights into the discriminative power of individual features. Glucose demonstrated the clearest separation between the two classes, with diabetic patients exhibiting a distinctly higher distribution of plasma glucose concentrations than non-diabetic patients. The distribution of non-diabetic patients is centred around 110 mg/dL, while the diabetic distribution is shifted rightward and centred around 140 mg/dL. BMI and Age also showed meaningful separation between the two classes, though with greater overlap than Glucose. The Pregnancies feature showed a moderate positive skew in both classes, while Insulin and SkinThickness exhibited heavily right-skewed distributions in both classes, consistent with the large number of imputed values."),
      figureCaption("Figure 4.3: Feature Distributions by Outcome Class (Green = Non-Diabetic, Red = Diabetic)"),
      body("The Pearson correlation heatmap (Figure 4.4) quantified the linear relationships between all pairs of features and between each feature and the Outcome variable. Glucose exhibited the strongest positive correlation with Outcome (r = 0.49), confirming its status as the most discriminative predictor of diabetes in this dataset — consistent with the well-established clinical role of plasma glucose concentration in diabetes diagnosis. BMI showed the second strongest correlation with Outcome (r = 0.29), followed by Age (r = 0.24), Pregnancies (r = 0.22), and DiabetesPedigreeFunction (r = 0.17). BloodPressure and SkinThickness showed weaker correlations with the Outcome variable (r = 0.07 and r = 0.07, respectively). Among the predictor features, a moderate positive correlation was observed between SkinThickness and BMI (r = 0.39) and between Age and Pregnancies (r = 0.54), which may introduce some degree of multicollinearity — a consideration of greater relevance for Logistic Regression than for tree-based algorithms."),
      figureCaption("Figure 4.4: Pearson Correlation Heatmap of All Dataset Features"),
      body("The boxplots (Figure 4.5) provided a complementary view of the class-conditional distributions and confirmed the patterns observed in the histograms. The boxplot for Glucose showed a clear upward shift in the median and interquartile range for diabetic patients relative to non-diabetic patients, with minimal overlap between the two distributions. Similar patterns of elevated medians for diabetic patients were observed for BMI and Age. All features, particularly Insulin, exhibited substantial outliers, visible as individual data points beyond the whiskers of the boxplots. These outliers were retained in the dataset rather than removed, as they represent genuine physiological variation among the patients."),
      figureCaption("Figure 4.5: Boxplots of Feature Distributions Grouped by Outcome Class"),

      sectionHeading("4.5 Model Training"),
      body("All three classification models were trained using the Scikit-learn machine learning library within a unified training script (train.py). The training script implements the complete pipeline: loading and preprocessing the data via the preprocess_pipeline function, instantiating the three classifiers with their respective hyperparameter configurations, fitting each classifier on the scaled training data, evaluating each classifier on the training, validation, and test sets, and serialising the trained models to disk."),
      body("Logistic Regression was instantiated using Scikit-learn's LogisticRegression class with max_iter=1000 (to ensure convergence of the iterative optimisation algorithm on this dataset), random_state=42 (for reproducibility), and all remaining hyperparameters set to their default values, including L2 regularisation with a regularisation strength parameter C=1.0. The L2 penalty helps prevent overfitting by penalising large coefficient values, contributing to the model's strong generalisation performance on the test set."),
      body("The Decision Tree was instantiated using DecisionTreeClassifier with random_state=42 and all other hyperparameters at their default values, including no restriction on maximum tree depth (max_depth=None). This configuration allows the tree to grow until all leaf nodes are pure (contain samples from only one class), resulting in a tree that perfectly memorises the training data — a textbook example of overfitting that is reflected in the training accuracy of 100%."),
      body("The Random Forest was instantiated using RandomForestClassifier with n_estimators=100 (an ensemble of 100 decision trees), random_state=42, and default values for all other hyperparameters. The use of bagging and random feature subsampling provides regularisation that reduces the variance of the ensemble relative to a single unpruned tree, though without explicit hyperparameter tuning, some overfitting is still evident in the results."),

      sectionHeading("4.6 Model Persistence"),
      body("Following training and evaluation, all three fitted model objects and the fitted StandardScaler were serialised to disk using the Joblib library. Joblib's dump function was used to write each Python object as a binary .pkl file to the models/ directory. The four files created are: logistic_regression.pkl, decision_tree.pkl, random_forest.pkl, and scaler.pkl."),
      body("Model persistence is an essential component of the system architecture for two principal reasons. First, it decouples the model training phase from the web application inference phase, enabling the computationally expensive training process to be executed once offline while allowing the lightweight web application to load pre-trained models at startup without incurring training overhead at prediction time. Second, it ensures that the exact scaler transformation fitted on the training data is consistently applied to new input data at inference time. Without persisting and reusing the same scaler, the web application would be unable to reproduce the exact normalisation applied during training, leading to distributional mismatch and degraded prediction accuracy."),

      sectionHeading("4.7 Web Application Development"),
      subHeading("4.7.1 Application Architecture"),
      body("The web application was developed using the Flask micro web framework and follows the Model-View-Controller (MVC) architectural pattern. In this architecture, the trained Scikit-learn model objects and the fitted StandardScaler constitute the Model layer, responsible for the core predictive logic. The Jinja2 HTML templates rendered by Flask constitute the View layer, responsible for presenting the user interface and displaying prediction results. The Flask route handlers defined in app.py constitute the Controller layer, responsible for receiving HTTP requests, orchestrating the interaction between the Model and View layers, and returning HTTP responses."),
      body("The application consists of a single route (/) that handles both GET requests (serving the empty prediction form to the user) and POST requests (receiving submitted form data, performing prediction, and returning the result). The application is structured within the app/ directory, with the main application module in app/app.py, Jinja2 HTML templates in app/templates/, and static assets (CSS stylesheets) in app/static/. The trained model files are stored in the models/ directory, which is located one level above the app/ directory in the project root."),

      subHeading("4.7.2 User Interface"),
      body("The user interface was designed to be clean, intuitive, and accessible to users without technical expertise in machine learning. The application presents a single-page web form (illustrated in Figure 4.6) containing eight numerical input fields corresponding to the eight clinical features used by the prediction models: Pregnancies, Glucose (mg/dL), Blood Pressure (mmHg), Skin Thickness (mm), Insulin (μU/mL), BMI (kg/m²), Diabetes Pedigree Function, and Age. Each input field is clearly labelled with the feature name and unit of measurement, and includes a placeholder text displaying the plausible range of values for that feature, derived from the dataset, to guide the user in entering realistic values."),
      body("A dropdown menu allows the user to select which of the three trained models — Logistic Regression, Decision Tree, or Random Forest — to use for the prediction, enabling direct comparison of the predictions produced by different algorithms for the same patient data. A prominently displayed Predict button submits the form data as an HTTP POST request to the server."),
      body("Following form submission, the result section (Figure 4.7) is rendered below the form, displaying the prediction label (Diabetic or Not Diabetic), the estimated probability of diabetes as a percentage, and the name of the model used. A brief advisory message is also displayed: a caution to seek professional medical advice for a positive prediction, or a reassuring note to maintain a healthy lifestyle for a negative prediction. The result section is colour-coded — red for a positive (Diabetic) prediction and green for a negative (Not Diabetic) prediction — to provide immediate visual feedback. Client-side HTML5 form validation ensures that all fields are completed with numeric values before the form is submitted."),
      figureCaption("Figure 4.6: Flask Web Application — Patient Data Input Form"),
      figureCaption("Figure 4.7: Flask Web Application — Prediction Result Display"),

      subHeading("4.7.3 Prediction Pipeline"),
      body("Upon receipt of a POST request, the Flask route handler executes the following prediction pipeline. First, the eight feature values submitted via the HTML form are extracted from the request object and converted to floating-point numbers. If any conversion fails due to non-numeric input, an error message is returned to the user. Second, the extracted values are assembled into a NumPy array with shape (1, 8), representing a single patient record. Third, this array is passed to the transform method of the pre-loaded StandardScaler object to apply the same z-score normalisation that was fitted on the training data. Fourth, the scaled array is passed to the predict method of the selected pre-loaded model to obtain the predicted class label (0 or 1), and to the predict_proba method to obtain the estimated class probabilities. Fifth, the predicted label is mapped to a human-readable string (Diabetic or Not Diabetic), the diabetes probability is extracted from the positive class probability and multiplied by 100 to express it as a percentage, and these results are passed to the Jinja2 template for rendering."),

      // ════════════════════════════════════════════════════════════════
      // CHAPTER 5
      // ════════════════════════════════════════════════════════════════
      chapterHeading("CHAPTER FIVE: RESULTS AND DISCUSSION"),
      sectionHeading("5.1 Introduction"),
      body("This chapter presents and analyses the results obtained from the experimental evaluation of the three machine learning classification models developed in this project. It begins with a summary of the key findings from the Exploratory Data Analysis. It then presents the performance metrics of each model on the validation set, followed by the definitive performance metrics on the held-out test set. A comparative analysis of the three models is then conducted, followed by a broader discussion that situates the findings within the context of the existing literature. The chapter concludes with an honest assessment of the limitations of this study."),

      sectionHeading("5.2 Exploratory Data Analysis Results"),
      body("The EDA revealed several important characteristics of the Pima Indians Diabetes Dataset that directly informed the preprocessing strategy and provided a basis for interpreting the model results. The dataset exhibits a moderate class imbalance, with non-diabetic cases constituting 65.1% and diabetic cases 34.9% of all records. This imbalance means that a trivial classifier that always predicts the majority class (non-diabetic) would achieve a baseline accuracy of 65.1%, providing an important context for evaluating the performance of the trained models."),
      body("Glucose emerged as the single most predictive feature, exhibiting the strongest positive Pearson correlation with the Outcome variable (r = 0.49) and the clearest visual separation between diabetic and non-diabetic distributions in both the histograms and boxplots. This finding is clinically consistent with the established diagnostic role of plasma glucose concentration in diabetes diagnosis and is well-corroborated in the existing literature. BMI (r = 0.29) and Age (r = 0.24) were identified as the next most predictive features. BloodPressure and SkinThickness showed the weakest correlations with the outcome, suggesting they are the least discriminative features in this dataset."),
      body("The missing value analysis confirmed that Insulin (48.7% missing) and SkinThickness (29.6% missing) are the most data-sparse features in the dataset, with BloodPressure, BMI, and Glucose also affected to lesser degrees. The decision to replace these missing values with column-specific medians rather than means was validated by the highly skewed distribution of Insulin values, which would render the mean a poor representative of the central tendency. The EDA thus provided a rigorous empirical basis for all preprocessing decisions."),

      sectionHeading("5.3 Model Performance on Validation Set"),
      body("Table 5.1 presents the performance metrics of all three classifiers evaluated on the validation set, which was used during the model development phase to compare models and guide the selection of the best candidate for final evaluation."),
      ...perfTable("Table 5.1: Model Performance on Validation Set", [
        ["Logistic Regression", "71.30%", "59.46%", "55.00%", "57.14%"],
        ["Decision Tree",       "72.17%", "59.52%", "62.50%", "60.98%"],
        ["Random Forest",       "74.78%", "65.71%", "57.50%", "61.33%"],
      ]),
      body("On the validation set, Random Forest achieved the highest accuracy (74.78%) and precision (65.71%), followed by Decision Tree (72.17%) and Logistic Regression (71.30%). Decision Tree achieved the highest recall (62.50%) among the three models, while Logistic Regression recorded the lowest performance across all four metrics on the validation set. These validation results suggested that the tree-based models, and particularly Random Forest, might be the superior performers — a conclusion that the test set results subsequently challenged, underscoring the importance of reserving a separate test set for final model evaluation and not relying solely on validation set results for model selection."),

      sectionHeading("5.4 Model Performance on Test Set"),
      body("Table 5.2 presents the definitive performance metrics of all three classifiers evaluated on the held-out test set, which had no influence on any model development or selection decision. These results represent the most reliable estimate of each model's generalisation performance on unseen data."),
      ...perfTable("Table 5.2: Model Performance on Test Set", [
        ["Logistic Regression", "77.59%", "77.78%", "51.22%", "61.76%"],
        ["Random Forest",       "72.41%", "65.52%", "46.34%", "54.29%"],
        ["Decision Tree",       "67.24%", "54.55%", "43.90%", "48.65%"],
      ]),
      body("On the test set, Logistic Regression achieved the highest accuracy (77.59%), the highest precision (77.78%), and the highest F1-Score (61.76%), clearly outperforming both tree-based models. Random Forest ranked second with a test accuracy of 72.41% and precision of 65.52%, while Decision Tree ranked last with a test accuracy of 67.24% and the lowest precision (54.55%) and recall (43.90%) of the three models."),
      body("A particularly striking and instructive observation is the pronounced discrepancy between the training and test set accuracies for the Decision Tree and Random Forest models. Both tree-based classifiers achieved perfect training accuracy of 100% — a clear indication of overfitting — while their test accuracies were 67.24% and 72.41% respectively, representing substantial drops of 32.76 and 27.59 percentage points from their training performance. Logistic Regression, by contrast, achieved a training accuracy of 78.58% and a test accuracy of 77.59%, a marginal difference of less than one percentage point, demonstrating excellent generalisation and the absence of significant overfitting."),

      sectionHeading("5.5 Comparative Analysis"),
      body("Table 5.3 provides a consolidated summary of all three models across training, validation, and test sets, enabling a holistic comparison of their performance and generalisation characteristics."),
      tableCaption("Table 5.3: Comparative Summary of Model Performance Across All Data Splits"),
      new Table({
        width: { size: 9000, type: WidthType.DXA },
        columnWidths: [2600, 2133, 2133, 2134],
        rows: [
          new TableRow({ children: [hCell("Model", 2600), hCell("Train Accuracy", 2133), hCell("Val Accuracy", 2133), hCell("Test Accuracy", 2134)] }),
          new TableRow({ children: [dCellShaded("Logistic Regression", 2600), dCell("78.58%", 2133, AlignmentType.CENTER), dCell("71.30%", 2133, AlignmentType.CENTER), dCell("77.59%", 2134, AlignmentType.CENTER)] }),
          new TableRow({ children: [dCellShaded("Random Forest", 2600), dCell("100.00%", 2133, AlignmentType.CENTER), dCell("74.78%", 2133, AlignmentType.CENTER), dCell("72.41%", 2134, AlignmentType.CENTER)] }),
          new TableRow({ children: [dCellShaded("Decision Tree", 2600), dCell("100.00%", 2133, AlignmentType.CENTER), dCell("72.17%", 2133, AlignmentType.CENTER), dCell("67.24%", 2134, AlignmentType.CENTER)] }),
        ],
      }),
      emptyLine(),
      body("The comparative analysis clearly identifies Logistic Regression as the best-performing and most generalisable model for this dataset and task. Several factors account for this outcome. First, the PIDD is a relatively small dataset of 768 records, which is insufficient to support the complex, high-variance models that tree-based algorithms — particularly unpruned Decision Trees and large Random Forest ensembles — effectively require. On small datasets, simpler models with fewer parameters tend to generalise better, as the risk of memorising noise in the training data is lower. Second, the strong linear separability of the feature space, evidenced by the high correlation of Glucose and other features with the outcome, means that the linear decision boundary of Logistic Regression is well-suited to the problem. Third, the L2 regularisation inherent in the default Scikit-learn Logistic Regression implementation provides an additional layer of protection against overfitting that is absent in the default Decision Tree configuration."),
      body("The perfect training accuracy of the Decision Tree and Random Forest models is a classical manifestation of overfitting. An unpruned decision tree will always achieve 100% training accuracy, as it can grow sufficiently deep to memorise every training sample, including noise. While the bagging and random feature subsampling mechanisms of Random Forest provide meaningful variance reduction, they are insufficient to fully counteract overfitting without hyperparameter tuning — specifically, constraining the maximum depth of individual trees or specifying a minimum number of samples required at leaf nodes."),

      sectionHeading("5.6 Discussion"),
      body("The findings of this study are consistent with, and corroborated by, the broader diabetes prediction literature. Most notably, the test set accuracy of 77.59% achieved by Logistic Regression in this study is in precise agreement with the result of 77.6% reported by Zou et al. (2018) in their comparative study of Logistic Regression, ANN, and Decision Tree on the PIDD, despite differences in the specific preprocessing and evaluation procedures. This convergence of results across independent studies provides strong evidence for the robustness and reproducibility of Logistic Regression as a classifier for this dataset."),
      body("The underperformance of the Random Forest model in this study (72.41% test accuracy) relative to the results reported by Maniruzzaman et al. (2020) (81.97% accuracy) is readily explained by methodological differences. Maniruzzaman et al. (2020) employed an explicit feature selection procedure and performed hyperparameter optimisation, both of which are known to substantially improve the performance of Random Forest on small, noisy datasets by reducing the effective model complexity. The results of the present study suggest that, without these additional procedures, Random Forest cannot fully leverage its ensemble properties on a dataset as small as the PIDD and is outperformed by the simpler Logistic Regression baseline."),
      body("From a clinical perspective, the Recall value of 51.22% achieved by Logistic Regression, while the highest among the three models, reveals a significant limitation: the model fails to identify approximately half of all truly diabetic patients in the test set, classifying them as non-diabetic (false negatives). In a clinical screening context, this level of false negative rate would be unacceptable as a standalone diagnostic tool. However, it is important to contextualise this result within the limitations of the dataset: the small sample size, the significant missing data in key predictive features (particularly Insulin), and the absence of class imbalance remediation (such as SMOTE) all constrain the achievable recall. The system is intended to serve as a supportive screening tool to assist — rather than replace — clinical judgement, and in this context, a test accuracy of 77.59% with appropriate uncertainty communication (via probability estimates) represents a meaningful contribution."),
      body("The development of the Flask web application constitutes a practical contribution that extends the impact of this project beyond a purely academic exercise. By providing a functional, accessible interface for non-technical users to interact with the trained model, the application demonstrates a viable pathway for translating research prototypes into clinical decision-support tools. The ability to select between models within the interface adds educational value, enabling users to observe and compare the predictions of different algorithms for the same patient profile."),

      sectionHeading("5.7 Limitations"),
      body("This study is subject to several limitations that should be acknowledged and considered when interpreting its findings and assessing the generalisability of its conclusions."),
      body("First, the dataset used is small, comprising only 768 patient records — a sample size that is insufficient for training high-capacity machine learning models and that limits the statistical power of the comparative evaluation. Modern machine learning research typically requires tens of thousands to millions of labelled examples to achieve state-of-the-art performance. The small dataset size is the primary reason for the overfitting observed in the tree-based models and may also account for the variance in performance metrics across the different data splits."),
      body("Second, the dataset is demographically homogeneous, containing records exclusively from female patients of Pima Indian heritage aged 21 years and above. This demographic restriction severely limits the generalisability of the trained models to broader populations, including males, younger individuals, and individuals of different ethnic backgrounds. A model trained on this dataset may exhibit significant performance degradation when applied to patients outside the original demographic group."),
      body("Third, class imbalance — with non-diabetic cases representing 65.1% of the dataset — was not explicitly addressed through resampling techniques such as SMOTE (Synthetic Minority Oversampling Technique) or cost-sensitive learning. This imbalance is a contributing factor to the relatively low recall values observed across all three models, as classifiers trained on imbalanced datasets tend to be biased towards the majority class."),
      body("Fourth, no systematic hyperparameter optimisation was performed on any of the three models. Decision Tree and Random Forest in particular are known to be highly sensitive to hyperparameter choices, and the default configurations employed in this study result in significant overfitting. Techniques such as grid search cross-validation or Bayesian optimisation could substantially improve the performance of these models."),
      body("Fifth, the web application developed in this study is a research prototype and has not undergone clinical validation, regulatory evaluation, or integration testing with real-world electronic health record systems. It should not be used for actual clinical diagnosis without appropriate validation and regulatory approval."),

      // ════════════════════════════════════════════════════════════════
      // CHAPTER 6
      // ════════════════════════════════════════════════════════════════
      chapterHeading("CHAPTER SIX: CONCLUSION AND FUTURE WORK"),
      sectionHeading("6.1 Conclusion"),
      body("This project set out to develop, evaluate, and compare machine learning classification models for predicting the likelihood of diabetes mellitus using the Pima Indians Diabetes Dataset, and to deploy the best-performing model within an accessible web-based application. In pursuit of this aim, a comprehensive and methodologically rigorous pipeline was implemented, encompassing data acquisition, missing value imputation, feature scaling, exploratory data analysis, model training, multi-metric evaluation, and web application deployment."),
      body("Three supervised machine learning classifiers were trained and evaluated: Logistic Regression, Decision Tree, and Random Forest. The experimental results demonstrated that Logistic Regression was the best-performing model, achieving a test set accuracy of 77.59%, precision of 77.78%, recall of 51.22%, and F1-Score of 61.76%. Both the Decision Tree and Random Forest classifiers exhibited significant overfitting — characterised by perfect training accuracy (100%) and substantially lower test accuracy (67.24% and 72.41% respectively) — attributable to the small size of the dataset and the absence of hyperparameter regularisation. These findings underscore the principle that model complexity should be calibrated to the size and characteristics of the available training data, and that simpler, well-regularised models frequently outperform more complex counterparts on small, real-world datasets."),
      body("The data preprocessing phase, particularly the identification and imputation of biologically implausible zero values representing missing data, was identified as the most critical determinant of model performance — consistent with the findings of Maniruzzaman et al. (2020). The EDA confirmed that Glucose is the single most discriminative feature for diabetes prediction in this dataset, with a Pearson correlation of 0.49 with the outcome variable, followed by BMI and Age. These findings are clinically intuitive and well-supported by the existing medical literature."),
      body("The development of a Flask-based web application represents a meaningful practical contribution, demonstrating a viable and reproducible approach to deploying machine learning models as interactive clinical decision-support tools accessible to non-technical users. The application allows users to input patient health parameters, select a prediction model, and receive a real-time prediction with an associated probability estimate, bridging the gap between research prototype and practical application."),
      body("In conclusion, this project provides empirical evidence that machine learning — and Logistic Regression in particular — can serve as a valuable, data-driven tool to support the early detection of diabetes mellitus. While the models developed are not intended to replace clinical diagnosis, they demonstrate the potential of accessible, low-cost computational approaches to extend diagnostic capability in resource-limited healthcare settings and to support more timely and equitable access to diabetes screening."),

      sectionHeading("6.2 Recommendations"),
      body("Based on the findings and limitations of this study, the following recommendations are offered for researchers and practitioners seeking to build upon this work:"),
      bullet("The acquisition of larger, more diverse, and demographically representative datasets is strongly recommended. Datasets that include patients of both sexes, a wide age range, and multiple ethnicities would substantially improve the generalisability and clinical applicability of the trained models."),
      bullet("Systematic hyperparameter optimisation using techniques such as k-fold cross-validated grid search (GridSearchCV) or randomised search should be applied to the Decision Tree and Random Forest classifiers. Constraining the maximum tree depth, minimum samples per leaf, and minimum samples per split are particularly impactful hyperparameters for controlling overfitting in tree-based models."),
      bullet("Class imbalance remediation techniques, including Synthetic Minority Oversampling Technique (SMOTE), class-weighted loss functions, or threshold optimisation, should be applied to improve the recall of the positive (diabetic) class — the most clinically critical performance metric."),
      bullet("Feature selection or dimensionality reduction techniques, such as Recursive Feature Elimination (RFE), Principal Component Analysis (PCA), or LASSO regularisation, should be explored to identify the most informative subset of features and potentially improve model generalisation."),
      bullet("The evaluation framework should be extended to include additional metrics such as the Area Under the Receiver Operating Characteristic Curve (AUC-ROC) and calibration plots, which provide a more comprehensive assessment of binary classifier performance, particularly in the context of medical decision-making."),

      sectionHeading("6.3 Future Work"),
      body("Several promising directions for future work are identified, building upon the foundation established in this project:"),
      bullet("Deep learning architectures, including Multi-Layer Perceptrons (MLPs) and, for larger datasets, Convolutional Neural Networks applied to tabular data, could be explored to determine whether the increased representational capacity of neural networks translates to improved diabetes prediction performance."),
      bullet("The integration of Explainable AI (XAI) techniques — such as SHAP (SHapley Additive exPlanations) values and LIME (Local Interpretable Model-agnostic Explanations) — would provide clinicians with feature-level insights into individual predictions, significantly enhancing the transparency, trustworthiness, and clinical utility of the prediction system."),
      bullet("The web application could be extended into a mobile application using frameworks such as React Native or Flutter, enabling deployment on smartphones and broadening accessibility to patients and healthcare workers in community settings."),
      bullet("Integration with real-time health monitoring data from wearable devices — such as continuous glucose monitors, smartwatches, and fitness trackers — would enable dynamic, longitudinal diabetes risk monitoring rather than single-point prediction."),
      bullet("Integration of the prediction system with hospital information systems (HIS) or electronic health record (EHR) platforms would facilitate its adoption as a practical clinical decision-support tool within existing healthcare workflows."),
      bullet("Future studies should also explore multi-class classification tasks, such as distinguishing between pre-diabetes, Type 1, and Type 2 diabetes, as well as predicting the severity of diabetic complications, to broaden the clinical utility of machine learning-based diabetes assessment systems."),

      // ════════════════════════════════════════════════════════════════
      // REFERENCES
      // ════════════════════════════════════════════════════════════════
      pageBreak(),
      chapterHeading("REFERENCES"),
      body("American Diabetes Association (2022). Standards of Medical Care in Diabetes — 2022. Diabetes Care, 45(Supplement 1), pp. S1–S264. https://doi.org/10.2337/dc22-Sint"),
      body("Breiman, L. (2001). Random Forests. Machine Learning, 45(1), pp. 5–32. https://doi.org/10.1023/A:1010933404324"),
      body("IDF — International Diabetes Federation (2021). IDF Diabetes Atlas. 10th edition. Brussels: International Diabetes Federation. Available at: https://www.diabetesatlas.org [Accessed: 20 January 2026]."),
      body("Islam, M.M.F., Ferdousi, R., Rahman, S. and Bushra, H.Y. (2020). Likelihood Prediction of Diabetes at Early Stage Using Data Mining Techniques. In: Gupta, M., Konar, D., Bhattacharyya, S. and Biswas, S. (eds.) Computer Vision and Machine Intelligence in Medical Image Analysis. Advances in Intelligent Systems and Computing, vol. 992. Singapore: Springer, pp. 113–125. https://doi.org/10.1007/978-981-13-8798-2_12"),
      body("Kavakiotis, I., Tsave, O., Salifoglou, A., Maglaveras, N., Vlahavas, I. and Chouvarda, I. (2017). Machine Learning and Data Mining Methods in Diabetes Research. Computational and Structural Biotechnology Journal, 15, pp. 104–116. https://doi.org/10.1016/j.csbj.2016.12.005"),
      body("Maniruzzaman, M., Rahman, M.J., Al-MehediHasan, M., Suri, H.S., Abedin, M.M., El-Baz, A. and Suri, J.S. (2020). Accurate Diabetes Risk Stratification Using Machine Learning: Role of Missing Value and Outliers. Journal of Medical Systems, 44(3), p. 92. https://doi.org/10.1007/s10916-019-1499-4"),
      body("McKinney, W. (2010). Data Structures for Statistical Computing in Python. Proceedings of the 9th Python in Science Conference (SciPy 2010), Austin, Texas, pp. 56–61. https://doi.org/10.25080/Majora-92bf1922-00a"),
      body("Mitchell, T.M. (1997). Machine Learning. New York: McGraw-Hill."),
      body("Obermeyer, Z. and Emanuel, E.J. (2016). Predicting the Future — Big Data, Machine Learning, and Clinical Medicine. New England Journal of Medicine, 375(13), pp. 1216–1219. https://doi.org/10.1056/NEJMp1606181"),
      body("Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., Blondel, M., Prettenhofer, P., Weiss, R., Dubourg, V., Vanderplas, J., Passos, A., Cournapeau, D., Brucher, M., Perrot, M. and Duchesnay, E. (2011). Scikit-learn: Machine Learning in Python. Journal of Machine Learning Research, 12, pp. 2825–2830."),
      body("Sarwar, M.A., Kamal, N., Hamid, W. and Shah, M.A. (2018). Prediction of Diabetes Using Machine Learning Algorithms in Healthcare. Proceedings of the 24th International Conference on Automation and Computing (ICAC 2018), Newcastle upon Tyne, UK, pp. 1–6. https://doi.org/10.23919/IConAC.2018.8749050"),
      body("Sisodia, D. and Sisodia, D.S. (2018). Prediction of Diabetes using Classification Algorithms. Procedia Computer Science, 132, pp. 1578–1585. https://doi.org/10.1016/j.procs.2018.05.122"),
      body("Smith, J.W., Everhart, J.E., Dickson, W.C., Knowler, W.C. and Johannes, R.S. (1988). Using the ADAP Learning Algorithm to Forecast the Onset of Diabetes Mellitus. Proceedings of the Annual Symposium on Computer Application in Medical Care, Washington DC, pp. 261–265."),
      body("Topol, E.J. (2019). High-performance medicine: the convergence of human and artificial intelligence. Nature Medicine, 25(1), pp. 44–56. https://doi.org/10.1038/s41591-018-0300-7"),
      body("WHO — World Health Organization (2022). Diabetes. Fact Sheet. Geneva: World Health Organization. Available at: https://www.who.int/news-room/fact-sheets/detail/diabetes [Accessed: 15 January 2026]."),
      body("Zou, Q., Qu, K., Luo, Y., Yin, D., Ju, Y. and Tang, H. (2018). Predicting Diabetes Mellitus With Machine Learning Techniques. Frontiers in Genetics, 9, p. 515. https://doi.org/10.3389/fgene.2018.00515"),

      // ════════════════════════════════════════════════════════════════
      // APPENDIX
      // ════════════════════════════════════════════════════════════════
      pageBreak(),
      chapterHeading("APPENDIX A: CODE LISTINGS"),
      body("The following appendix contains the complete source code for the three primary Python modules developed in this project."),

      sectionHeading("Appendix A1: preprocess.py — Data Preprocessing Module"),
      new Paragraph({
        spacing: { before: 80, after: 200 },
        children: [new TextRun({
          text:
`import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
import joblib
import os

ZERO_MEANS_MISSING = ['Glucose', 'BloodPressure', 'SkinThickness', 'Insulin', 'BMI']

def load_data(filepath):
    df = pd.read_csv(filepath)
    return df

def replace_zeros_with_median(df):
    df = df.copy()
    for col in ZERO_MEANS_MISSING:
        median = df[col].replace(0, np.nan).median()
        df[col] = df[col].replace(0, median)
    return df

def split_data(df):
    X = df.drop('Outcome', axis=1)
    y = df['Outcome']
    X_train, X_temp, y_train, y_temp = train_test_split(
        X, y, test_size=0.30, random_state=42, stratify=y)
    X_val, X_test, y_val, y_test = train_test_split(
        X_temp, y_temp, test_size=0.50, random_state=42, stratify=y_temp)
    return X_train, X_val, X_test, y_train, y_val, y_test

def scale_features(X_train, X_val, X_test, scaler_path=None):
    scaler = StandardScaler()
    X_train_scaled = scaler.fit_transform(X_train)
    X_val_scaled   = scaler.transform(X_val)
    X_test_scaled  = scaler.transform(X_test)
    if scaler_path:
        os.makedirs(os.path.dirname(scaler_path), exist_ok=True)
        joblib.dump(scaler, scaler_path)
    return X_train_scaled, X_val_scaled, X_test_scaled, scaler

def preprocess_pipeline(filepath, scaler_path=None):
    df = load_data(filepath)
    df = replace_zeros_with_median(df)
    X_train, X_val, X_test, y_train, y_val, y_test = split_data(df)
    X_train_s, X_val_s, X_test_s, scaler = scale_features(
        X_train, X_val, X_test, scaler_path)
    return X_train_s, X_val_s, X_test_s, y_train, y_val, y_test, scaler`,
          font: "Courier New", size: 18,
        })],
      }),

      sectionHeading("Appendix A2: train.py — Model Training and Evaluation"),
      new Paragraph({
        spacing: { before: 80, after: 200 },
        children: [new TextRun({
          text:
`import os, joblib
from sklearn.linear_model  import LogisticRegression
from sklearn.tree          import DecisionTreeClassifier
from sklearn.ensemble      import RandomForestClassifier
from sklearn.metrics       import (accuracy_score, precision_score,
                                   recall_score, f1_score)
from preprocess import preprocess_pipeline

DATA_PATH  = os.path.join('data', 'diabetes.csv')
MODELS_DIR = 'models'
SCALER_PATH = os.path.join(MODELS_DIR, 'scaler.pkl')

def get_models():
    return {
        'Logistic Regression': LogisticRegression(max_iter=1000, random_state=42),
        'Decision Tree':       DecisionTreeClassifier(random_state=42),
        'Random Forest':       RandomForestClassifier(n_estimators=100, random_state=42),
    }

def evaluate(model, X, y, split_name):
    y_pred = model.predict(X)
    acc  = accuracy_score(y, y_pred)
    prec = precision_score(y, y_pred, zero_division=0)
    rec  = recall_score(y, y_pred, zero_division=0)
    f1   = f1_score(y, y_pred, zero_division=0)
    print(f"  [{split_name}] Acc:{acc:.4f} Prec:{prec:.4f} Rec:{rec:.4f} F1:{f1:.4f}")
    return {'accuracy':acc,'precision':prec,'recall':rec,'f1':f1}

def train_and_evaluate():
    X_train,X_val,X_test,y_train,y_val,y_test,_ = preprocess_pipeline(
        DATA_PATH, scaler_path=SCALER_PATH)
    os.makedirs(MODELS_DIR, exist_ok=True)
    models  = get_models()
    results = {}
    for name, model in models.items():
        print(f"--- {name} ---")
        model.fit(X_train, y_train)
        results[name] = {
            'train': evaluate(model, X_train, y_train, 'Train'),
            'val':   evaluate(model, X_val,   y_val,   'Val'),
            'test':  evaluate(model, X_test,  y_test,  'Test'),
        }
        joblib.dump(model, os.path.join(MODELS_DIR,
            f"{name.lower().replace(' ','_')}.pkl"))
    return results

if __name__ == '__main__':
    train_and_evaluate()`,
          font: "Courier New", size: 18,
        })],
      }),

      sectionHeading("Appendix A3: app/app.py — Flask Web Application"),
      new Paragraph({
        spacing: { before: 80, after: 200 },
        children: [new TextRun({
          text:
`import os, sys
import numpy as np
import joblib
from flask import Flask, render_template, request

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
app = Flask(__name__)
MODELS_DIR = os.path.join(os.path.dirname(__file__), '..', 'models')
FEATURE_NAMES = ['Pregnancies','Glucose','BloodPressure','SkinThickness',
                 'Insulin','BMI','DiabetesPedigreeFunction','Age']

def load_artifacts():
    scaler = joblib.load(os.path.join(MODELS_DIR, 'scaler.pkl'))
    models = {
        'Logistic Regression': joblib.load(
            os.path.join(MODELS_DIR,'logistic_regression.pkl')),
        'Decision Tree':       joblib.load(
            os.path.join(MODELS_DIR,'decision_tree.pkl')),
        'Random Forest':       joblib.load(
            os.path.join(MODELS_DIR,'random_forest.pkl')),
    }
    return scaler, models

@app.route('/', methods=['GET','POST'])
def index():
    prediction, probability, selected_model = None, None, None
    form_values, error = {}, None
    scaler, models = load_artifacts()
    if request.method == 'POST':
        try:
            selected_model = request.form.get('model','Random Forest')
            form_values = {f: request.form.get(f,'') for f in FEATURE_NAMES}
            features = np.array([[float(form_values[f]) for f in FEATURE_NAMES]])
            features_scaled = scaler.transform(features)
            model = models[selected_model]
            pred  = model.predict(features_scaled)[0]
            prob  = model.predict_proba(features_scaled)[0][1] * 100
            prediction  = 'Diabetic' if pred == 1 else 'Not Diabetic'
            probability = round(prob, 1)
        except ValueError:
            error = "Please fill in all fields with valid numbers."
    return render_template('index.html',
        prediction=prediction, probability=probability,
        selected_model=selected_model, model_names=list(models.keys()),
        feature_names=FEATURE_NAMES, form_values=form_values, error=error)

if __name__ == '__main__':
    app.run(debug=True)`,
          font: "Courier New", size: 18,
        })],
      }),

    ],
  }],
});

const outputPath = path.resolve(process.argv[2] || 'Diabetes_Project_Report.docx');
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outputPath, buf);
  console.log(`Report saved to ${outputPath}`);
}).catch(error => {
  console.error('Report generation failed:', error.message);
  process.exitCode = 1;
});
