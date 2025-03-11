const { promises: fs } = require('fs')
const path = require('path')
const os = require('os')
const mammoth = require('mammoth')
const puppeteer = require('puppeteer')

/**
 * Convert DOCX buffer to PDF with flexible return options
 * @param {Buffer} fileBuffer - Buffer of the DOCX file
 * @param {string} originalFilename - Original filename (for output naming)
 * @param {Object} options - Options for return type and output path
 * @param {string} [options.returnType='buffer'] - Return type: 'buffer', 'file', or 'base64'
 * @param {string} [options.outputDir=os.tmpdir()] - Output directory for saved files
 * @returns {Promise<{buffer: Buffer, filename: string, base64?: string}>} - PDF Buffer, filename, and optionally base64 string
 */
async function convertDocxToPdf(fileBuffer, originalFilename, options = {}) {
  const { returnType = 'buffer', outputDir = os.tmpdir() } = options
  
  try {
    // Create temporary directory
    const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'doc-convert-'))

    // Write DOCX file to temp directory
    const tempFilePath = path.join(tempDir, `input.docx`)
    await fs.writeFile(tempFilePath, fileBuffer)

    // Convert DOCX to HTML using Mammoth
    const { value: htmlContent } = await mammoth.convertToHtml({ path: tempFilePath })

    // Create temporary HTML file
    const tempHtmlPath = path.join(tempDir, 'output.html')
    await fs.writeFile(tempHtmlPath, htmlContent)

    // Convert HTML to PDF using Puppeteer
    const outputPath = path.join(outputDir, originalFilename.replace(/\.docx$/i, '.pdf'))
    const browser = await puppeteer.launch({
      headless: true,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox'
      ]
    })
    
    const page = await browser.newPage()
    await page.setContent(htmlContent, {
      waitUntil: 'networkidle0'
    })
    await page.pdf({ path: outputPath, format: 'A4', printBackground: true })
    await browser.close()

    // Read converted PDF
    const pdfBuffer = await fs.readFile(outputPath)
    
    // Generate Base64 string if requested
    let base64 = null
    if (returnType === 'base64') {
      base64 = pdfBuffer.toString('base64')
    }
    
    // Return according to specified type
    if (returnType === 'buffer') {
      return {
        buffer: pdfBuffer,
        filename: path.basename(outputPath)
      }
    } else if (returnType === 'file') {
      return {
        filename: outputPath
      }
    } else if (returnType === 'base64') {
      return {
        base64: base64,
        filename: path.basename(outputPath)
      }
    } else {
      throw new Error(`Invalid returnType specified: ${returnType}`)
    }
  } catch (error) {
    console.error('Error converting DOCX to PDF:', error)
    throw new Error('Failed to convert DOCX document to PDF')
  }
}

module.exports = { convertDocxToPdf }
