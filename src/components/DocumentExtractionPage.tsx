import { useState } from 'react'

const extractionApiUrl = 'https://coa-webapp-etfmfkd3cceph2a8.canadacentral-01.azurewebsites.net/api/documents/extract-test-results'

const systemPrompt = `You are a precise document data extraction engine specialized in Certificates of Analysis (CoA) for chemical products.
Your task is to extract ONLY the "Test Results" table from the provided document content and return it as strict, valid JSON.
The uploaded PDF may contain one or more pages and may include documents other than a Certificate of Analysis.
If multiple COA pages exist, combine all extracted laboratory test results into a single JSON array.

The Test Results table in these documents follows this structure:
"Inspection characteristic/-method" — this cell contains TWO lines: the characteristic/test name on the first line, and a method/standard code on the second line (e.g., "QM-AA-634", "DIN 53715", "QM-AA-146").
- "Specification" — the acceptance criteria for the test, shown as free text. This may be blank, a single-sided limit with an operator (e.g., "<= 1.9", ">= 98.0"), or a range (e.g., "268 - 275").
- "Result" — the actual measured/reported value, shown as free text exactly as printed. This may be a plain number, a number with a unit suffix (e.g., "0.4 %", "274 °C"), or qualitative text (e.g., "Equal to standard").

Extraction rules:
1. Split the "Inspection characteristic/-method" cell into two fields: "characteristic" (the test name) and "method" (the method/standard code on the line below it). Set "method" to null if no code is present.
2. Capture the "Specification" column exactly as printed, as a single string, in the "specification" field. Set to null if blank/empty.
3. Capture the "Result" column exactly as printed, as a single string, in the "result" field, preserving any unit symbol shown (e.g., "%", "°C"). Do not split it into separate numeric/unit fields.
4. Treat "-", blank cells, or "N/A" as null, not as a string.
5. Do not infer, calculate, normalize, or fabricate any values not explicitly present in the document.
6. Preserve the original order of tests as they appear in the document.
7. If the document spans multiple pages, merge all Test Results rows from every page into a single "test_results" array, in reading order. Ignore all non-table content — header/address blocks, order/material/batch details, boilerplate/legal/footer text (e.g., shelf life notes, EN 10204-3.1 statement, signature disclaimer).
8. Do not include explanation, commentary, markdown, or code fences — output raw JSON only.

Return the result in exactly this JSON schema:

{
  "test_results": [
    {
      "characteristic": string or null,
      "method": string or null,
      "specification": string or null,
      "result": string or null
    }
  ]
}

Output nothing except this JSON object.`

const userPrompt = `Extract the Test Results table from the following Certificate of Analysis document and return the JSON as specified in your instructions.

The supplier and column format for this document are unknown or unspecified; detect the table structure directly from the content and normalize it according to your defined rules.

If the Document Content below is empty, missing, or unreadable, return the schema with all fields set to null and an empty "test_results" array — do not fabricate any data.`

const fileToBase64 = (selectedFile: File): Promise<string> => new Promise((resolve, reject) => {
  const reader = new FileReader()
  reader.onload = () => {
    const result = reader.result
    if (typeof result !== 'string') {
      reject(new Error('Unable to read the selected file.'))
      return
    }
    resolve(result.split(',')[1] || result)
  }
  reader.onerror = () => reject(new Error('Unable to read the selected file.'))
  reader.readAsDataURL(selectedFile)
})

export function DocumentExtractionPage({ onBack }: { onBack: () => void }) {
  const [file, setFile] = useState<File | null>(null)
  const [response, setResponse] = useState<unknown>(null)
  const [error, setError] = useState('')
  const [isSubmitting, setIsSubmitting] = useState(false)

  const submitFile = async () => {
    if (!file || isSubmitting) return

    setIsSubmitting(true)
    setResponse(null)
    setError('')

    try {
      const fileContent = await fileToBase64(file)
      const result = await fetch(extractionApiUrl, {
        method: 'POST',
        headers: {
          accept: 'application/json',
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          fileContent,
          fileName: file.name,
          systemPropmt: systemPrompt,
          userPrompt,
        }),
      })
      const contentType = result.headers.get('content-type') || ''
      const resultBody = contentType.includes('application/json') ? await result.json() : await result.text()

      if (!result.ok) {
        const errorMessage = typeof resultBody === 'string' ? resultBody : JSON.stringify(resultBody)
        throw new Error(errorMessage || `Request failed with status ${result.status}.`)
      }

      setResponse(resultBody)
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : 'Unable to complete the API request.')
    } finally {
      setIsSubmitting(false)
    }
  }

  const htmlResponse = getHtmlResponse(response)

  return <main className="page-content new-submission-page">
    <div className="page-heading">
      <p>Document Extraction</p>
      <span>Today is {new Date().toLocaleDateString('en-GB', { weekday: 'long', day: '2-digit', month: 'short', year: 'numeric' })}</span>
    </div>

    <section className="form-section document-extraction-form">
      <h2>Upload Document</h2>
      <div className="form-grid-2">
        <label className="form-field">
          <span className="form-label">File</span>
          <input type="file" onChange={(event) => { setFile(event.target.files?.[0] || null); setResponse(null); setError('') }} />
        </label>
      </div>
      {file && <p className="selected-file">Selected file: {file.name}</p>}
      {error && <div className="form-message document-error" role="alert">{error}</div>}
      <div className="form-actions document-extraction-actions">
        <button className="primary-action" type="button" disabled={!file || isSubmitting} onClick={() => void submitFile()}>{isSubmitting ? 'Processing...' : 'Submit'}</button>
        <button type="button" disabled={isSubmitting} onClick={onBack}>Back</button>
      </div>
    </section>

    {response !== null && <section className="form-section document-response-section">
      <h2>API Response</h2>
      {htmlResponse ? <div className="document-response document-response-html" aria-live="polite" dangerouslySetInnerHTML={{ __html: htmlResponse }} /> : <pre className="document-response" aria-live="polite">{typeof response === 'string' ? response : JSON.stringify(response, null, 2)}</pre>}
    </section>}
  </main>
}

function getHtmlResponse(response: unknown): string {
  if (!response || typeof response !== 'object' || !('html' in response)) return ''
  const html = (response as { html?: unknown }).html
  return typeof html === 'string' && html.trim() ? sanitizeHtml(html) : ''
}

function sanitizeHtml(html: string): string {
  const documentFragment = new DOMParser().parseFromString(html, 'text/html')
  documentFragment.querySelectorAll('script, style, iframe, object, embed').forEach((node) => node.remove())
  documentFragment.querySelectorAll('*').forEach((element) => {
    Array.from(element.attributes).forEach((attribute) => {
      const attributeName = attribute.name.toLowerCase()
      if (attributeName.startsWith('on') || attributeName === 'srcdoc') element.removeAttribute(attribute.name)
      if ((attributeName === 'href' || attributeName === 'src') && /^(javascript|data):/i.test(attribute.value.trim())) {
        element.removeAttribute(attribute.name)
      }
    })
  })
  return documentFragment.body.innerHTML
}
