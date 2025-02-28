const fs = require('fs')
const path = require('path')
const { convertDocxToPdf } = require('./index')

async function test() {
  try {
    // Load a sample DOCX file
    const docxBuffer = fs.readFileSync(path.join(__dirname, 'sample.docx'))

    // Test 1: Return as Buffer (default)
    const bufferResult = await convertDocxToPdf(docxBuffer, 'sample.docx')
    fs.writeFileSync(path.join(__dirname, bufferResult.filename), bufferResult.buffer)
    console.log(`Buffer PDF saved as ${bufferResult.filename}`)

    // Test 2: Return as File
    const fileResult = await convertDocxToPdf(docxBuffer, 'sample.docx', { returnType: 'file' })
    console.log(`File PDF saved as ${fileResult.filename}`)

    // Test 3: Return as Base64
    const base64Result = await convertDocxToPdf(docxBuffer, 'sample.docx', { returnType: 'base64' })
    console.log(`Base64 string of PDF: ${base64Result.base64.slice(0, 30)}...`)
  } catch (error) {
    console.error('Test failed:', error)
  }
}

test()
