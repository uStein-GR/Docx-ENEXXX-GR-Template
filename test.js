import React from "react";
import { Document, Packer, Paragraph, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType } from "docx";
import { saveAs } from "file-saver";

export default function App() {
    const generateWordDocument = () => {
        const doc = new Document({
            sections: [
                {
                    properties: {},
                    children: [
                        // Title Section
                        new Paragraph({
                            text: "King Mongkut’s University of Technology Thonburi",
                            heading: HeadingLevel.HEADING_1,
                            alignment: AlignmentType.CENTER,
                        }),
                        new Paragraph({
                            text: "Department of Electronics and Telecommunication Engineering",
                            heading: HeadingLevel.HEADING_2,
                            alignment: AlignmentType.CENTER,
                        }),
                        new Paragraph({
                            text: "Course Portfolio",
                            heading: HeadingLevel.HEADING_2,
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 300 },
                        }),

                        // Course Code and Instructor
                        new Paragraph({
                            text: "ENEXXX -------------------------------- 1/2024",
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 100 },
                        }),
                        new Paragraph({
                            text: "Instructor: --------------------------------",
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 400 },
                        }),

                        // Course Learning Outcomes (CLOs)
                        new Paragraph({
                            text: "Course Learning Outcomes (CLOs)",
                            heading: HeadingLevel.HEADING_3,
                            spacing: { before: 300, after: 200 },
                        }),
                        new Paragraph({
                            text: "At the end of the course, students should be able to:",
                            spacing: { after: 200 },
                        }),

                        // Student Outcomes (SOs) and PIs
                        new Paragraph({
                            text: "Student Outcomes (SOs)",
                            heading: HeadingLevel.HEADING_3,
                            spacing: { before: 300, after: 100 },
                        }),
                        new Paragraph({ text: "PI XX", spacing: { after: 100 } }),
                        new Paragraph({ text: "X", spacing: { after: 200 } }),

                        // 1. Methods to Assess the CLOs
                        new Paragraph({
                            text: "1. Methods to Assess the CLOs",
                            heading: HeadingLevel.HEADING_3,
                            spacing: { before: 300, after: 200 },
                        }),
                        new Table({
                            rows: [
                                new TableRow({
                                    children: [
                                        new TableCell({
                                            children: [new Paragraph("CLO")],
                                            width: { size: 20, type: WidthType.PERCENTAGE },
                                        }),
                                        new TableCell({
                                            children: [new Paragraph("Method of assessment")],
                                            width: { size: 40, type: WidthType.PERCENTAGE },
                                        }),
                                        new TableCell({
                                            children: [new Paragraph("Assessment tool")],
                                            width: { size: 20, type: WidthType.PERCENTAGE },
                                        }),
                                        new TableCell({
                                            children: [new Paragraph("Criteria for the indicators")],
                                            width: { size: 20, type: WidthType.PERCENTAGE },
                                        }),
                                    ],
                                }),
                                // Add more rows here for each PI if needed
                            ],
                        }),

                        // 2. Result of CLOs Assessment
                        new Paragraph({
                            text: "2. Result of CLOs Assessment",
                            heading: HeadingLevel.HEADING_3,
                            spacing: { before: 300, after: 200 },
                        }),
                        new Paragraph({
                            text: "Direct Assessment",
                            spacing: { after: 200 },
                        }),
                        new Paragraph({
                            text: "The target is to have at least 60% of the students achieve each performance indicator in Level 4 or 5. The column “Result” indicates whether the accumulated percentage meets the target.",
                            spacing: { after: 200 },
                        }),
                        new Paragraph({
                            text: "Sample size = number of students enrolled in the course = XX",
                            spacing: { after: 200 },
                        }),
                        new Table({
                            rows: [
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("CLO")] }),
                                        new TableCell({ children: [new Paragraph("Average skill level")] }),
                                        new TableCell({ children: [new Paragraph("Distribution of skill level")] }),
                                        new TableCell({ children: [new Paragraph("Cumulation at levels ≥ 4")] }),
                                        new TableCell({ children: [new Paragraph("Meet target?")] }),
                                    ],
                                }),
                                // Additional rows for assessment data
                            ],
                        }),

                        // Self-Evaluation Section
                        new Paragraph({
                            text: "3. Self-Evaluation on the Validity and Reliability of the Direct Assessment",
                            heading: HeadingLevel.HEADING_3,
                            spacing: { before: 300, after: 200 },
                        }),
                        new Paragraph({
                            text: "Using the Department-issued 5-scaled rubric on validity and reliability, the instructor evaluated the validity and reliability of the CLO assessment as follows:",
                            spacing: { after: 200 },
                        }),

                        // Continuous Quality Improvement
                        new Paragraph({
                            text: "4. Continuous Quality Improvement",
                            heading: HeadingLevel.HEADING_3,
                            spacing: { before: 300, after: 200 },
                        }),
                        new Paragraph({ text: "Faculty Evaluation of Attainment of CLOs", spacing: { after: 100 } }),
                        new Paragraph({ text: "Student Evaluation of the Course Strengths and Weaknesses", spacing: { after: 100 } }),
                        new Paragraph({
                            text: "Remedy Plan",
                            spacing: { after: 100 },
                        }),
                        new Paragraph({
                            text: "The course instructor proposed the following actions as possible remedies.",
                            spacing: { after: 200 },
                        }),

                        // Appendices
                        new Paragraph({
                            text: "Appendix A: Embedded Questions Done by a Top Student",
                            heading: HeadingLevel.HEADING_4,
                            spacing: { before: 300, after: 100 },
                        }),
                        new Paragraph({
                            text: "Appendix B: Embedded Questions Done by a Median Student",
                            heading: HeadingLevel.HEADING_4,
                            spacing: { after: 100 },
                        }),
                        new Paragraph({
                            text: "Appendix C: Embedded Questions Done by a Bottom Student",
                            heading: HeadingLevel.HEADING_4,
                            spacing: { after: 100 },
                        }),
                    ],
                },
            ],
        });

        Packer.toBlob(doc)
            .then((blob) => {
                saveAs(blob, "course_portfolio_template_complete.docx");
            })
            .catch((error) => console.error("Document generation error:", error));
    };

    return (
        <div style={{ padding: "20px" }}>
            <h1>Word Document Template</h1>
            <button onClick={generateWordDocument}>Generate Word Document</button>
        </div>
    );
}
