# DOCX to PDF Converter

Convert DOCX files to PDF using Mammoth and Puppeteer.

## Installation

```sh
npm install docx-pdf-converter
```


### Usage 

```js

const fs = require('fs')
const { convertDocxToPdf } = require('docx-pdf-converter')

async function convert() {
  const docxBuffer = fs.readFileSync('path/to/input.docx')
  
  // Return as Buffer (default)
  const result = await convertDocxToPdf(docxBuffer, 'input.docx')
  fs.writeFileSync(result.filename, result.buffer)

  // Return as File
  const fileResult = await convertDocxToPdf(docxBuffer, 'input.docx', { returnType: 'file' })
  console.log(`File saved as ${fileResult.filename}`)

  // Return as Base64
  const base64Result = await convertDocxToPdf(docxBuffer, 'input.docx', { returnType: 'base64' })
  console.log(`Base64 string: ${base64Result.base64.slice(0, 30)}...`)
}

convert()

```
