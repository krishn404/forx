// src/FormatterForm.jsx
import React, { useState } from 'react';
import { Document, Packer, Paragraph, PageBreak, AlignmentType, HeadingLevel, TextRun } from 'docx';
import { saveAs } from 'file-saver';

const FormatterForm = () => {
    const [font, setFont] = useState('Times New Roman');
    const [fontSizeMain, setFontSizeMain] = useState(20); // Font size in points
    const [fontSizeHeading, setFontSizeHeading] = useState(24); // Font size in points
    const [lineSpacing, setLineSpacing] = useState(1.5); // Line spacing in lines
    const [margin, setMargin] = useState(1440); // Margin in twips (1 inch = 1440 twips)
    const [pageNumbering, setPageNumbering] = useState('bottom-center');
    const [length, setLength] = useState(5); // Number of pages

    const generateDocx = () => {
        // Create paragraphs for the document
        const paragraphs = [
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Title Page",
                        bold: true,
                        size: fontSizeHeading * 2, // docx uses half-points
                        font: font,
                    }),
                ],
                heading: HeadingLevel.TITLE,
                alignment: AlignmentType.CENTER,
                spacing: { after: 1000 },
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Main text here.",
                        size: fontSizeMain * 2, // docx uses half-points
                        font: font,
                    }),
                ],
                alignment: AlignmentType.LEFT,
                spacing: { line: lineSpacing * 240 }, // 240 half-points per line
            }),
        ];

        // Add page breaks and page numbers dynamically
        for (let i = 0; i < length; i++) {
            paragraphs.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `Page ${i + 1}`,
                            size: fontSizeMain * 2, // docx uses half-points
                            font: font,
                        }),
                    ],
                    alignment: pageNumbering === 'bottom-center' ? AlignmentType.CENTER : AlignmentType.RIGHT,
                    spacing: { before: 100, after: 100 },
                }),
            );
            if (i < length - 1) {
                paragraphs.push(new PageBreak());
            }
        }

        // Create document
        const doc = new Document({
            styles: {
                default: {
                    document: {
                        run: {
                            font: font,
                        },
                    },
                    paragraph: {
                        spacing: {
                            line: lineSpacing * 240, // 240 half-points per line
                        },
                        run: {
                            font: font,
                        },
                    },
                },
            },
            sections: [
                {
                    properties: {
                        page: {
                            margin: {
                                top: margin,
                                right: margin,
                                bottom: margin,
                                left: margin,
                            },
                        },
                    },
                    children: paragraphs,
                },
            ],
        });

        Packer.toBlob(doc).then((blob) => {
            saveAs(blob, "formatted-document.docx");
        });
    };

    return (
        <div>
            <h1>Document Formatter</h1>
            <form>
                <div>
                    <label>Font:</label>
                    <input type="text" value={font} onChange={(e) => setFont(e.target.value)} />
                </div>
                <div>
                    <label>Main Font Size (points):</label>
                    <input type="number" value={fontSizeMain} onChange={(e) => setFontSizeMain(Number(e.target.value))} />
                </div>
                <div>
                    <label>Heading Font Size (points):</label>
                    <input type="number" value={fontSizeHeading} onChange={(e) => setFontSizeHeading(Number(e.target.value))} />
                </div>
                <div>
                    <label>Line Spacing (lines):</label>
                    <input type="number" step="0.1" value={lineSpacing} onChange={(e) => setLineSpacing(Number(e.target.value))} />
                </div>
                <div>
                    <label>Margin (in twips):</label>
                    <input type="number" value={margin} onChange={(e) => setMargin(Number(e.target.value))} />
                </div>
                <div>
                    <label>Page Numbering:</label>
                    <select value={pageNumbering} onChange={(e) => setPageNumbering(e.target.value)}>
                        <option value="bottom-center">Bottom Center</option>
                        <option value="bottom-right">Bottom Right</option>
                    </select>
                </div>
                <div>
                    <label>Length (pages):</label>
                    <input type="number" value={length} onChange={(e) => setLength(Number(e.target.value))} />
                </div>
                <button type="button" onClick={generateDocx}>Generate DOCX</button>
            </form>
        </div>
    );
};

export default FormatterForm;
