    Webserver with low-cost frontends interface for export data into word files that fix template, which use for SeniorProject 2024

<Initialize Settings by following commands>
- Asusual Create the React Application
    npx create-react-app docx-test

- Install DOCX into package.json
    npm install docx file-saver

- Install the last version of DOCX
    npm install docx@7.1.0
    npm install docx@latest file-saver@latest

<Testing the initial system by add following command into file App.js in subdirectory file src>

import React from "react";
import { Document, Packer, Paragraph } from "docx";
import { saveAs } from "file-saver";

export default function App() {
    const generateWordDocument = () => {
        const doc = new Document({
            sections: [
                {
                    children: [
                        new Paragraph({
                            text: "This is a minimal Word document.",
                            heading: "Heading1"
                        })
                    ]
                }
            ]
        });

        Packer.toBlob(doc)
            .then((blob) => {
                saveAs(blob, "minimal_test.docx");
            })
            .catch((error) => console.error("Document generation error:", error));
    };

    return (
        <div style={{ padding: "20px" }}>
            <h1>Word Document Test</h1>
            <button onClick={generateWordDocument}>Generate Word Document</button>
        </div>
    );
}

-
