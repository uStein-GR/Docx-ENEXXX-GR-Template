## Initialize Settings by following commands
- Asusual Create the React Application
a    npx create-react-app docx-test

- Install DOCX into package.json
b    npm install docx file-saver

- Install the last version of DOCX
c    npm install docx@7.1.0
c    npm install docx@latest file-saver@latest

## Testing the initial system by add following command into file App.js in subdirectory file src

```
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
```