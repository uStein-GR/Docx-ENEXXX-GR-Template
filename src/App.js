import React from "react";
import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  WidthType,
  TextRun,
  AlignmentType,
} from "docx";
import { saveAs } from "file-saver";

export default function ExportWordDoc() {
  const generateDoc = () => {
    const doc = new Document({
      sections: [
        {
          // ตั้งค่าขอบหน้า 1" (1440 twip) รอบเอกสาร
          properties: {
            page: {
              margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
            },
          },
          children: [
            // --- Header ---
            new Paragraph({
              children: [
                new TextRun({
                  text: "King Mongkut’s University of Technology Thonburi",
                  bold: true,
                  size: 24,
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Department of Electronics and Telecommunication Engineering",
                  size: 24,
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Course Portfolio",
                  bold: true,
                  size: 24,
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { after: 300 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "ENEXXX \t--------------------------------\t1/2024",
                  size: 24,
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Instructor: \t\t--------------------------------",
                  size: 24,
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { after: 400 },
            }),

            // --- Course Learning Outcomes ---
            new Paragraph({
              children: [
                new TextRun({
                  text: "Course Learning Outcomes (CLOs)",
                  bold: true,
                  size: 24,
                }),
              ],
              spacing: { before: 200, after: 100 },
            }),

            // --- Combined CLO + SO Table (70/30) ---
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
                insideH: 100,
                insideV: 100,
              },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 70, type: WidthType.PERCENTAGE },
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text:
                                "At the end of the course, students should be able to:",
                              italics: true,
                              size: 24,
                            }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      width: { size: 30, type: WidthType.PERCENTAGE },
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: "Student Outcomes\n(SOs)",
                              italics: true,
                              size: 24,
                            }),
                          ],
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                  ],
                }),
                ...[1, 2, 3].map(() =>
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [new TextRun({ text: "PI XX", size: 24 })],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [new TextRun({ text: "X", size: 24 })],
                          }),
                        ],
                      }),
                    ],
                  })
                ),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "(PI = Performance indicator)",
                  italics: true,
                  size: 24,
                }),
              ],
              spacing: { before: 100, after: 300 },
            }),

            // --- 1. Methods to Assess the CLOs ---
            new Paragraph({
              children: [
                new TextRun({
                  text: "1. Methods to Assess the CLOs",
                  bold: true,
                  size: 24,
                }),
              ],
              spacing: { before: 200, after: 100 },
            }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
                insideH: 100,
                insideV: 100,
              },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 15, type: WidthType.PERCENTAGE },
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({ text: "CLO", size: 24 }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      width: { size: 35, type: WidthType.PERCENTAGE },
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: "Method of assessment",
                              size: 24,
                            }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      width: { size: 15, type: WidthType.PERCENTAGE },
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({ text: "Assessment tool", size: 24 }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      width: { size: 35, type: WidthType.PERCENTAGE },
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text:
                                "Criteria for the indicators",
                              size: 24,
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
                ...[1, 2, 3].flatMap(() => [
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({ text: "PI X", size: 24 }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({
                                text:
                                  "Direct: Embedded test question pertaining to the PI",
                                size: 24,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [new TextRun({ text: "", size: 24 })],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({
                                text:
                                  "My own 5-scaled rubric that matches the Department’s scale (e.g., 5=excellent, …, 1=very poor)",
                                size: 24,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [new TextRun({ text: "", size: 24 })],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [new TextRun({ text: "Indirect:", size: 24 })],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [new TextRun({ text: "N/A", size: 24 })],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [new TextRun({ text: "N/A", size: 24 })],
                          }),
                        ],
                      }),
                    ],
                  }),
                ]),
              ],
            }),

            // --- 2. Result of CLOs Assessment ---
            new Paragraph({
              children: [
                new TextRun({
                  text: "2. Result of CLOs Assessment",
                  bold: true,
                  size: 24,
                }),
              ],
              spacing: { before: 300, after: 100 },
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Direct Assessment", bold: true, size: 24 }),
              ],
              spacing: { after: 100 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text:
                    "The target is to have at least 60% of the students achieve each performance indicator in Level 4 or 5. The column “Result” indicates whether the accumulated percentage meets the target.",
                  size: 24,
                }),
              ],
              spacing: { after: 100 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text:
                    "Sample size = number of students enrolled in the course = XX",
                  size: 24,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
                insideH: 100,
                insideV: 100,
              },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      rowSpan: 2,
                      children: [
                        new Paragraph({
                          children: [new TextRun({ text: "CLO", size: 24 })],
                        }),
                      ],
                    }),
                    new TableCell({
                      rowSpan: 2,
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({ text: "Average skill level", size: 24 }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      columnSpan: 5,
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({ text: "Distribution of skill level", size: 24 }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      rowSpan: 2,
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({ text: "Cumulation\nat levels ≥ 4", size: 24 }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      rowSpan: 2,
                      children: [
                        new Paragraph({
                          children: [new TextRun({ text: "Meet target?", size: 24 })],
                        }),
                      ],
                    }),
                  ],
                }),
                new TableRow({
                  children: [5, 4, 3, 2, 1].map((n) =>
                    new TableCell({
                      children: [
                        new Paragraph({ children: [new TextRun({ text: `${n}`, size: 24 })] }),
                      ],
                    })
                  ),
                }),
                ...[1, 2, 3].map(() =>
                  new TableRow({
                    children: Array(9)
                      .fill(null)
                      .map(() =>
                        new TableCell({
                          children: [
                            new Paragraph({ children: [new TextRun({ text: "", size: 24 })] }),
                          ],
                        })
                      ),
                  })
                ),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text:
                    "The work of the top, median, and bottom students in terms of their skill levels are shown in Appendices A, B, and C, respectively. The top student had PIs XX of 5, 5, and 5, respectively. The median student had PIs XX of 3, 4-, and -3 respectively. The bottom student had PI XX of 1, 2, and 1, respectively.",
                  size: 24,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Indirect Assessment", bold: true, size: 24 }),
              ],
            }),
            new Paragraph({
              children: [new TextRun({ text: "N/A", size: 24 })],
              spacing: { after: 300 },
            }),

            // --- 3. Self-Evaluation ---
            new Paragraph({
              children: [
                new TextRun({
                  text:
                    "3. Self-Evaluation on the Validity and Reliability of the Direct Assessment",
                  bold: true,
                  size: 24,
                }),
              ],
              spacing: { before: 300, after: 100 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text:
                    "Using the Department-issued 5-scaled rubric on validity and reliability, the instructor evaluated the validity and reliability of the CLO assessment as follows:",
                  size: 24,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
                insideH: 100,
                insideV: 100,
              },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 15, type: WidthType.PERCENTAGE }, children:
                      [new Paragraph({ children: [new TextRun({ text: "SO", size: 24 })] })],
                    }),
                    new TableCell({
                      width: { size: 60, type: WidthType.PERCENTAGE }, children:
                      [new Paragraph({ children: [new TextRun({ text: "Parameter of the assessment tool", size: 24 })] })],
                    }),
                    new TableCell({
                      width: { size: 25, type: WidthType.PERCENTAGE }, children:
                      [new Paragraph({ children: [new TextRun({ text: "Level (Meaning)", size: 24 })] })],
                    }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "X", size: 24 })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Validity", size: 24 })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 24 })] })] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 24 })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Reliability", size: 24 })] })] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 24 })] })] }),
                  ],
                }),
              ],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Justification of the specified levels", size: 24 })],
              spacing: { after: 300 },
            }),

            // --- 4. Continuous Quality Improvement ---
            new Paragraph({
              children: [
                new TextRun({
                  text: "4. Continuous Quality Improvement",
                  bold: true,
                  size: 24,
                }),
              ],
              spacing: { before: 200, after: 100 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Faculty Evaluation of Attainment of CLOs",
                  bold: true,
                  size: 24,
                }),
              ],
              spacing: { after: 100 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Student Evaluation of the Course Strengths and Weaknesses",
                  bold: true,
                  size: 24,
                }),
              ],
              spacing: { after: 100 },
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Remedy Plan", bold: true, size: 24 }),
              ],
              spacing: { after: 100 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text:
                    "The course instructor proposed the following actions as possible remedies.",
                  size: 24,
                }),
              ],
              spacing: { after: 300 },
            }),

            // --- Appendices ---
            new Paragraph({
              children: [
                new TextRun({ text: "Appendix A", bold: true, size: 24 }),
                new TextRun({ text: "\nEmbedded Questions Done by a Top Student", size: 24 }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Appendix B", bold: true, size: 24 }),
                new TextRun({ text: "\nEmbedded Questions Done by a Median Student", size: 24 }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Appendix C", bold: true, size: 24 }),
                new TextRun({ text: "\nEmbedded Questions Done by a Bottom Student", size: 24 }),
              ],
            }),
          ],
        },
      ],
    });

    Packer.toBlob(doc).then((blob) =>
      saveAs(blob, "ENEXXX_Course_Portfolio_Final.docx")
    );
  };

  return (
    <div style={{ padding: 20 }}>
      <h2>Export Official Course Portfolio</h2>
      <button onClick={generateDoc}>Download DOCX</button>
    </div>
  );
}