import { useEffect, useState } from 'react'
import { ArrowDownToLine, FileText, ImageIcon, Layers3, LoaderCircle, Sparkles, Upload, X } from 'lucide-react'
import './App.css'

const API_BASE_URL = (import.meta.env.VITE_API_URL || 'http://127.0.0.1:8000').replace(/\/$/, '')

const DEFAULT_TRANSCRIPT = `Paste a meeting transcript, workshop notes, or raw requirements here.

Example:
Client: Northstar Studio
Project: Launch documentation
We discussed the product goals, timeline, approval process, and the need for a clean PDF the team can share with stakeholders.`

function base64ToBlobUrl(base64Value) {
  const binaryString = window.atob(base64Value)
  const bytes = new Uint8Array(binaryString.length)
  for (let index = 0; index < binaryString.length; index += 1) {
    bytes[index] = binaryString.charCodeAt(index)
  }
  return URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }))
}

function textToBlobUrl(textValue) {
  return URL.createObjectURL(new Blob([textValue], { type: 'text/x-tex;charset=utf-8' }))
}

function base64ToDocxBlobUrl(base64Value) {
  const binaryString = window.atob(base64Value)
  const bytes = new Uint8Array(binaryString.length)
  for (let index = 0; index < binaryString.length; index += 1) {
    bytes[index] = binaryString.charCodeAt(index)
  }
  return URL.createObjectURL(
    new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }),
  )
}

function SectionCard({ title, icon, children }) {
  return (
    <section className="panel-card">
      <div className="panel-card__header">
        <span className="panel-card__icon">{icon}</span>
        <h2>{title}</h2>
      </div>
      {children}
    </section>
  )
}

export default function App() {
  const [sourceText, setSourceText] = useState(DEFAULT_TRANSCRIPT)
  const [agentInstructions, setAgentInstructions] = useState('')
  const [logoFile, setLogoFile] = useState(null)
  const [pdfFiles, setPdfFiles] = useState([])
  const [imageItems, setImageItems] = useState([])
  const [preview, setPreview] = useState(null)
  const [isGenerating, setIsGenerating] = useState(false)
  const [error, setError] = useState('')

  useEffect(() => {
    return () => {
      if (preview?.pdfUrl) URL.revokeObjectURL(preview.pdfUrl)
      if (preview?.texUrl) URL.revokeObjectURL(preview.texUrl)
      if (preview?.docxUrl) URL.revokeObjectURL(preview.docxUrl)
    }
  }, [preview?.pdfUrl, preview?.texUrl, preview?.docxUrl])

  const handleLogoChange = (event) => {
    const file = event.target.files?.[0] ?? null
    setLogoFile(file)
  }

  const handleImagesAdd = (event) => {
    const files = Array.from(event.target.files || [])
    const newItems = files.map((file) => ({
      file,
      description: '',
      previewUrl: URL.createObjectURL(file),
    }))
    setImageItems((prev) => [...prev, ...newItems])
    event.target.value = ''
  }

  const handleImageDescriptionChange = (index, value) => {
    setImageItems((prev) =>
      prev.map((item, i) => (i === index ? { ...item, description: value } : item)),
    )
  }

  const handleImageRemove = (index) => {
    setImageItems((prev) => {
      URL.revokeObjectURL(prev[index].previewUrl)
      return prev.filter((_, i) => i !== index)
    })
  }

  const handleGenerate = async (event) => {
    event.preventDefault()
    setError('')
    setIsGenerating(true)

    try {
      const payload = new FormData()
      payload.append('source_text', sourceText)
      payload.append('agent_instructions', agentInstructions)

      if (logoFile) {
        payload.append('logo', logoFile)
      }

      pdfFiles.forEach((file) => {
        payload.append('source_pdfs', file)
      })

      if (imageItems.length > 0) {
        payload.append('image_descriptions', JSON.stringify(imageItems.map((item) => item.description)))
        imageItems.forEach(({ file }) => payload.append('source_images', file))
      }

      const response = await fetch(`${API_BASE_URL}/api/documents/preview/`, {
        method: 'POST',
        body: payload,
      })

      const responseData = await response.json().catch(() => null)

      if (!response.ok) {
        const message =
          responseData?.errors
            ? JSON.stringify(responseData.errors)
            : 'Document generation failed.'
        throw new Error(message)
      }

      const pdfUrl = base64ToBlobUrl(responseData.pdf_base64)
      const texUrl = textToBlobUrl(responseData.latex_source || '')
      const docxUrl = base64ToDocxBlobUrl(responseData.docx_base64 || '')
      const texFilename = (responseData.filename || 'document.pdf').replace(/\.pdf$/i, '.tex')
      const docxFilename =
        responseData.docx_filename ||
        (responseData.filename || 'document.pdf').replace(/\.pdf$/i, '.docx')

      setPreview({
        pdfUrl,
        texUrl,
        docxUrl,
        texFilename,
        docxFilename,
        filename: responseData.filename,
        document: responseData.document,
        latexSource: responseData.latex_source || '',
        generation_mode: responseData.generation_mode || 'openai',
      })
    } catch (generationError) {
      setError(generationError.message)
    } finally {
      setIsGenerating(false)
    }
  }

  return (
    <div className="app-shell">
      <div className="app-backdrop app-backdrop--one" />
      <div className="app-backdrop app-backdrop--two" />

      <header className="hero">
        <div className="hero__copy">
          <p className="eyebrow">
            <Sparkles size={14} /> Automatic documentation creator
          </p>
          <h1>Turn a transcript into a polished client PDF.</h1>
          <p className="hero__description">
            Paste the full transcript or notes dump once, optionally attach a logo, and let OpenAI draft the LaTeX
            source and documentation text for the PDF preview.
          </p>
        </div>
      </header>

      <main className="workspace-grid">
        <section className="editor-column">
          <SectionCard title="Text Input" icon={<FileText size={18} />}>
            <form className="creator-form" onSubmit={handleGenerate}>
              <label className="field-group">
                <span className="field-label">Instruction for the agent (optional)</span>
                <input
                  className="creator-field"
                  value={agentInstructions}
                  placeholder='Example: "I want a document that explains the use of Celonis to a client."'
                  onChange={(event) => setAgentInstructions(event.target.value)}
                />
              </label>

              <label className="field-group">
                <span className="field-label">Add text or notes</span>
                <textarea
                  className="creator-field creator-textarea"
                  value={sourceText}
                  placeholder="Paste your transcript here"
                  onChange={(event) => setSourceText(event.target.value)}
                  rows={18}
                />
                <span className="field-note">
                  This is the only required input. The backend uses it to generate the title, sections, LaTeX, and
                  PDF.
                </span>
              </label>

              <label className="field-group field-group--file">
                <span className="field-label">
                  <Upload size={14} /> Reference PDFs (optional)
                </span>
                <input
                  className="creator-file"
                  type="file"
                  accept="application/pdf"
                  multiple
                  onChange={(event) => setPdfFiles(Array.from(event.target.files || []))}
                />
                <span className="field-note">
                  Upload one or more PDF documents. The agent will read them and incorporate their content.
                </span>
                {pdfFiles.length > 0 ? (
                  <div className="file-pills">
                    {pdfFiles.map((file) => (
                      <span key={file.name} className="file-pill">{file.name}</span>
                    ))}
                  </div>
                ) : null}
              </label>

              <div className="field-group field-group--file">
                <span className="field-label">
                  <ImageIcon size={14} /> Figures / Images (optional)
                </span>
                <input
                  className="creator-file"
                  type="file"
                  accept="image/png,image/jpeg,image/webp,image/gif"
                  multiple
                  onChange={handleImagesAdd}
                />
                <span className="field-note">
                  Upload images to include in the document. Add a description to help the agent place each one at the right spot and generate an appropriate caption.
                </span>
                {imageItems.length > 0 && (
                  <div className="figure-list">
                    {imageItems.map((item, index) => (
                      <div key={item.previewUrl} className="figure-item">
                        <img
                          className="figure-thumb"
                          src={item.previewUrl}
                          alt={`Figure ${index + 1}`}
                        />
                        <div className="figure-desc-wrap">
                          <span className="field-note">Figure {index + 1} — {item.file.name}</span>
                          <textarea
                            className="figure-desc-input"
                            rows={2}
                            placeholder="Describe what this image shows…"
                            value={item.description}
                            onChange={(e) => handleImageDescriptionChange(index, e.target.value)}
                          />
                        </div>
                        <button
                          type="button"
                          className="figure-remove-btn"
                          onClick={() => handleImageRemove(index)}
                          title="Remove image"
                        >
                          <X size={14} />
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <label className="field-group field-group--file">
                <span className="field-label">
                  <Upload size={14} /> Client logo, optional
                </span>
                <input className="creator-file" type="file" accept="image/png,image/jpeg,image/webp" onChange={handleLogoChange} />
                <span className="field-note">Attach a logo if the output needs branding.</span>
                {logoFile ? <span className="file-pill">{logoFile.name}</span> : null}
              </label>

              <div className="action-row">
                <button className="primary-button" type="submit" disabled={isGenerating}>
                  {isGenerating ? <LoaderCircle size={16} className="spin" /> : <ArrowDownToLine size={16} />}
                  {isGenerating ? 'Generating PDF...' : 'Generate preview'}
                </button>
              </div>
            </form>
          </SectionCard>
        </section>

        <section className="preview-column">
          <SectionCard title="PDF preview" icon={<Layers3 size={18} />}>
            {error ? <div className="alert alert--error">{error}</div> : null}

            {preview ? (
              <>
                <div className="preview-toolbar">
                  <div>
                    <p className="preview-label">Ready to download</p>
                    <h3>{preview.document.title}</h3>
                  </div>
                  <div className="download-actions">
                    <a className="secondary-button" href={preview.pdfUrl} download={preview.filename}>
                      <ArrowDownToLine size={16} /> Download PDF
                    </a>
                    <a className="secondary-button" href={preview.docxUrl} download={preview.docxFilename}>
                      <ArrowDownToLine size={16} /> Download Word
                    </a>
                    <a className="secondary-button" href={preview.texUrl} download={preview.texFilename}>
                      <ArrowDownToLine size={16} /> Download TEX
                    </a>
                  </div>
                </div>

                <div className="pdf-frame-wrap">
                  <iframe
                    className="pdf-frame"
                    src={preview.pdfUrl}
                    title="Generated document preview"
                  />
                </div>

                <details className="latex-panel">
                  <summary>View generated LaTeX source</summary>
                  <pre>{preview.latexSource}</pre>
                </details>
              </>
            ) : (
              <div className="empty-state">
                <div className="empty-state__icon">
                  <FileText size={24} />
                </div>
                <h3>No PDF generated yet</h3>
                <p>
                  Paste the transcript, optionally attach a logo, and generate the first preview to see the rendered
                  document here.
                </p>
              </div>
            )}
          </SectionCard>
        </section>
      </main>
    </div>
  )
}
