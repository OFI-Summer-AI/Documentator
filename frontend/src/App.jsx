import { useEffect, useState } from 'react'
import {
  ArrowDownToLine,
  ArrowUp,
  ArrowDown,
  CheckCircle2,
  ChevronRight,
  FileText,
  ImageIcon,
  Layers3,
  LoaderCircle,
  Pencil,
  Plus,
  RefreshCw,
  Sparkles,
  SlidersHorizontal,
  Trash2,
  Upload,
  X,
} from 'lucide-react'
import './App.css'

const _rawApiUrl = import.meta.env.VITE_API_URL || 'https://documentator-production.up.railway.app'
const API_BASE_URL = (_rawApiUrl.startsWith('http') ? _rawApiUrl : `https://${_rawApiUrl}`).replace(/\/$/, '')

const DEFAULT_TRANSCRIPT = `Paste a meeting transcript, workshop notes, or raw requirements here.

Example:
Client: Northstar Studio
Project: Launch documentation
We discussed the product goals, timeline, approval process, and the need for a clean PDF the team can share with stakeholders.`

const TEMPLATE_OPTIONS = [
  { value: 'general', label: 'General (agent decides structure)' },
  { value: 'meeting_minutes', label: 'Meeting Minutes' },
  { value: 'technical_spec', label: 'Technical Specification' },
  { value: 'proposal', label: 'Client Proposal' },
  { value: 'status_report', label: 'Status Report' },
]

const LENGTH_OPTIONS = [
  { value: 'brief', label: 'Brief' },
  { value: 'standard', label: 'Standard' },
  { value: 'detailed', label: 'Detailed' },
]

const TONE_OPTIONS = [
  { value: 'formal', label: 'Formal' },
  { value: 'consulting', label: 'Consulting' },
  { value: 'technical', label: 'Technical' },
  { value: 'casual', label: 'Casual' },
]

function base64ToBlobUrl(base64Value) {
  const binaryString = window.atob(base64Value)
  const bytes = new Uint8Array(binaryString.length)
  for (let index = 0; index < binaryString.length; index += 1) {
    bytes[index] = binaryString.charCodeAt(index)
  }
  return URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }))
}

function base64ToZipBlobUrl(base64Value) {
  const binaryString = window.atob(base64Value)
  const bytes = new Uint8Array(binaryString.length)
  for (let index = 0; index < binaryString.length; index += 1) {
    bytes[index] = binaryString.charCodeAt(index)
  }
  return URL.createObjectURL(new Blob([bytes], { type: 'application/zip' }))
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

let _sectionKeySeq = 0
function nextSectionKey() {
  _sectionKeySeq += 1
  return `section-${_sectionKeySeq}`
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

function CollapsibleSection({ title, icon, children, defaultOpen = false }) {
  const [isOpen, setIsOpen] = useState(defaultOpen)
  return (
    <div className="advanced-options-panel">
      <button
        type="button"
        className="advanced-options-summary"
        aria-expanded={isOpen}
        onClick={() => setIsOpen((prev) => !prev)}
      >
        {icon} {title}
        <ChevronRight size={14} className="advanced-options-chevron" />
      </button>
      {isOpen ? <div className="advanced-options-body">{children}</div> : null}
    </div>
  )
}

function TableContentEditor({ content, onChange }) {
  const headers = content?.headers || []
  const rows = content?.rows || []

  const updateHeader = (colIndex, value) => {
    const newHeaders = headers.map((h, i) => (i === colIndex ? value : h))
    onChange({ ...content, headers: newHeaders })
  }

  const updateCell = (rowIndex, colIndex, value) => {
    const newRows = rows.map((row, r) =>
      r === rowIndex ? row.map((cell, c) => (c === colIndex ? value : cell)) : row,
    )
    onChange({ ...content, rows: newRows })
  }

  const addRow = () => {
    onChange({ ...content, rows: [...rows, headers.map(() => '')] })
  }

  const removeRow = (rowIndex) => {
    onChange({ ...content, rows: rows.filter((_, r) => r !== rowIndex) })
  }

  return (
    <div className="table-editor">
      <div className="table-editor__row table-editor__row--header">
        {headers.map((h, colIndex) => (
          <input
            key={colIndex}
            className="table-editor__cell"
            value={h}
            onChange={(e) => updateHeader(colIndex, e.target.value)}
          />
        ))}
      </div>
      {rows.map((row, rowIndex) => (
        <div className="table-editor__row" key={rowIndex}>
          {row.map((cell, colIndex) => (
            <input
              key={colIndex}
              className="table-editor__cell"
              value={cell}
              onChange={(e) => updateCell(rowIndex, colIndex, e.target.value)}
            />
          ))}
          <button
            type="button"
            className="table-editor__remove-row"
            onClick={() => removeRow(rowIndex)}
            title="Remove row"
          >
            <X size={13} />
          </button>
        </div>
      ))}
      <button type="button" className="table-editor__add-row" onClick={addRow}>
        <Plus size={13} /> Add row
      </button>
    </div>
  )
}

function SectionEditorCard({ section, index, total, imageItems, isRegenerating, onChange, onDelete, onMove, onRegenerate }) {
  const isFigure = section.type === 'figure'

  return (
    <div className="section-editor-card">
      <div className="section-editor-card__header">
        {!isFigure ? (
          <input
            className="section-editor-card__title"
            value={section.title}
            placeholder="Section title"
            onChange={(e) => onChange(index, { title: e.target.value })}
          />
        ) : (
          <span className="section-editor-card__title section-editor-card__title--static">
            Figure {section.content?.figure_number ?? index + 1}
          </span>
        )}
        <span className="section-editor-card__type">{section.type}</span>
      </div>

      {section.type === 'paragraph' && (
        <textarea
          className="section-editor-card__textarea"
          rows={4}
          value={typeof section.content === 'string' ? section.content : ''}
          onChange={(e) => onChange(index, { content: e.target.value })}
        />
      )}

      {section.type === 'bullets' && (
        <textarea
          className="section-editor-card__textarea"
          rows={4}
          value={Array.isArray(section.content) ? section.content.join('\n') : ''}
          placeholder="One bullet per line"
          onChange={(e) => onChange(index, { content: e.target.value.split('\n') })}
        />
      )}

      {section.type === 'table' && (
        <TableContentEditor
          content={typeof section.content === 'object' ? section.content : { headers: [], rows: [] }}
          onChange={(newContent) => onChange(index, { content: newContent })}
        />
      )}

      {isFigure && (
        <div className="section-editor-card__figure">
          {imageItems[section.content?.figure_index]?.previewUrl ? (
            <img
              className="figure-thumb"
              src={imageItems[section.content.figure_index].previewUrl}
              alt={section.content?.caption || 'Figure'}
            />
          ) : null}
          <input
            className="section-editor-card__caption"
            value={section.content?.caption || ''}
            placeholder="Figure caption"
            onChange={(e) => onChange(index, { content: { ...section.content, caption: e.target.value } })}
          />
        </div>
      )}

      <div className="section-editor-card__actions">
        <button type="button" className="icon-button" title="Move up" disabled={index === 0} onClick={() => onMove(index, -1)}>
          <ArrowUp size={14} />
        </button>
        <button type="button" className="icon-button" title="Move down" disabled={index === total - 1} onClick={() => onMove(index, 1)}>
          <ArrowDown size={14} />
        </button>
        {!isFigure && (
          <button
            type="button"
            className="icon-button"
            title="Regenerate this section"
            disabled={isRegenerating}
            onClick={() => onRegenerate(index)}
          >
            {isRegenerating ? <LoaderCircle size={14} className="spin" /> : <RefreshCw size={14} />}
          </button>
        )}
        <button type="button" className="icon-button icon-button--danger" title="Delete section" onClick={() => onDelete(index)}>
          <Trash2 size={14} />
        </button>
      </div>
    </div>
  )
}

export default function App() {
  const [sourceText, setSourceText] = useState('')
  const [agentInstructions, setAgentInstructions] = useState('')
  const [logoFile, setLogoFile] = useState(null)
  const [pdfFiles, setPdfFiles] = useState([])
  const [imageItems, setImageItems] = useState([])

  const [template, setTemplate] = useState('general')
  const [targetLength, setTargetLength] = useState('standard')
  const [tone, setTone] = useState('formal')
  const [maxSections, setMaxSections] = useState('')
  const [reviewBeforeExport, setReviewBeforeExport] = useState(false)

  const [preview, setPreview] = useState(null)
  const [documentMeta, setDocumentMeta] = useState(null)
  const [sectionsDraft, setSectionsDraft] = useState(null)
  const [viewMode, setViewMode] = useState('empty') // 'empty' | 'review' | 'preview'
  const [regeneratingIndex, setRegeneratingIndex] = useState(null)

  const [isGenerating, setIsGenerating] = useState(false)
  const [isFinalizing, setIsFinalizing] = useState(false)
  const [error, setError] = useState('')

  useEffect(() => {
    return () => {
      if (preview?.pdfUrl) URL.revokeObjectURL(preview.pdfUrl)
      if (preview?.texZipUrl) URL.revokeObjectURL(preview.texZipUrl)
      if (preview?.docxUrl) URL.revokeObjectURL(preview.docxUrl)
    }
  }, [preview?.pdfUrl, preview?.texZipUrl, preview?.docxUrl])

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

  const appendGenerationOptions = (payload) => {
    payload.append('template', template)
    payload.append('target_length', targetLength)
    payload.append('tone', tone)
    if (maxSections) {
      payload.append('max_sections', maxSections)
    }
  }

  const appendSourceInputs = (payload) => {
    payload.append('source_text', sourceText)
    payload.append('agent_instructions', agentInstructions)

    pdfFiles.forEach((file) => {
      payload.append('source_pdfs', file)
    })

    if (imageItems.length > 0) {
      payload.append('image_descriptions', JSON.stringify(imageItems.map((item) => item.description)))
      imageItems.forEach(({ file }) => payload.append('source_images', file))
    }
  }

  const applyPreviewResponse = (responseData) => {
    const pdfUrl = base64ToBlobUrl(responseData.pdf_base64)
    const texZipUrl = base64ToZipBlobUrl(responseData.tex_zip_base64 || '')
    const docxUrl = base64ToDocxBlobUrl(responseData.docx_base64 || '')
    const texZipFilename = responseData.tex_zip_filename || 'document-latex.zip'
    const docxFilename =
      responseData.docx_filename ||
      (responseData.filename || 'document.pdf').replace(/\.pdf$/i, '.docx')

    setPreview({
      pdfUrl,
      texZipUrl,
      docxUrl,
      texZipFilename,
      docxFilename,
      filename: responseData.filename,
      document: responseData.document,
      latexSource: responseData.latex_source || '',
      generation_mode: responseData.generation_mode || 'openai',
    })
    setViewMode('preview')
  }

  const handleGenerate = async (event) => {
    event.preventDefault()
    setError('')

    if (reviewBeforeExport) {
      await handleGenerateSections()
      return
    }

    setIsGenerating(true)
    try {
      const payload = new FormData()
      appendSourceInputs(payload)
      appendGenerationOptions(payload)
      if (logoFile) {
        payload.append('logo', logoFile)
      }

      const response = await fetch(`${API_BASE_URL}/api/documents/preview/`, {
        method: 'POST',
        body: payload,
      })
      const responseData = await response.json().catch(() => null)

      if (!response.ok) {
        const message = responseData?.errors ? JSON.stringify(responseData.errors) : 'Document generation failed.'
        throw new Error(message)
      }

      applyPreviewResponse(responseData)
    } catch (generationError) {
      setError(generationError.message)
    } finally {
      setIsGenerating(false)
    }
  }

  const handleGenerateSections = async () => {
    setIsGenerating(true)
    setError('')
    try {
      const payload = new FormData()
      appendSourceInputs(payload)
      appendGenerationOptions(payload)

      const response = await fetch(`${API_BASE_URL}/api/documents/generate-sections/`, {
        method: 'POST',
        body: payload,
      })
      const responseData = await response.json().catch(() => null)

      if (!response.ok) {
        const message = responseData?.errors ? JSON.stringify(responseData.errors) : 'Section generation failed.'
        throw new Error(message)
      }

      setDocumentMeta({
        title: responseData.title,
        client_name: responseData.client_name,
        document_language: responseData.document_language,
        generation_mode: responseData.generation_mode,
        filename: responseData.filename,
      })
      setSectionsDraft(
        responseData.document_sections.map((section) => ({ ...section, _key: nextSectionKey() })),
      )
      setViewMode('review')
    } catch (generationError) {
      setError(generationError.message)
    } finally {
      setIsGenerating(false)
    }
  }

  const updateSection = (index, patch) => {
    setSectionsDraft((prev) => prev.map((s, i) => (i === index ? { ...s, ...patch } : s)))
  }

  const deleteSection = (index) => {
    setSectionsDraft((prev) => prev.filter((_, i) => i !== index))
  }

  const moveSection = (index, direction) => {
    setSectionsDraft((prev) => {
      const target = index + direction
      if (target < 0 || target >= prev.length) return prev
      const next = [...prev]
      ;[next[index], next[target]] = [next[target], next[index]]
      return next
    })
  }

  const regenerateSection = async (index) => {
    const section = sectionsDraft[index]
    setRegeneratingIndex(index)
    setError('')
    try {
      const response = await fetch(`${API_BASE_URL}/api/documents/regenerate-section/`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          source_text: sourceText,
          agent_instructions: agentInstructions,
          document_language: documentMeta?.document_language || 'en',
          template,
          target_length: targetLength,
          tone,
          section_title: section.title,
          section_type: section.type,
        }),
      })
      const responseData = await response.json().catch(() => null)
      if (!response.ok) {
        const message = responseData?.errors ? JSON.stringify(responseData.errors) : 'Section regeneration failed.'
        throw new Error(message)
      }
      updateSection(index, { title: responseData.section.title, content: responseData.section.content })
    } catch (regenerationError) {
      setError(regenerationError.message)
    } finally {
      setRegeneratingIndex(null)
    }
  }

  const handleFinalizeExport = async () => {
    setIsFinalizing(true)
    setError('')
    try {
      const cleanedSections = sectionsDraft.map((section) => {
        const clean = { title: section.title, type: section.type, content: section.content }
        if (clean.type === 'bullets' && Array.isArray(clean.content)) {
          clean.content = clean.content.map((c) => c.trim()).filter(Boolean)
        }
        return clean
      })

      const payload = new FormData()
      payload.append('title', documentMeta.title)
      payload.append('client_name', documentMeta.client_name || '')
      payload.append('document_language', documentMeta.document_language || 'en')
      payload.append('document_sections', JSON.stringify(cleanedSections))
      if (logoFile) {
        payload.append('logo', logoFile)
      }
      imageItems.forEach(({ file }) => payload.append('source_images', file))

      const response = await fetch(`${API_BASE_URL}/api/documents/render/`, {
        method: 'POST',
        body: payload,
      })
      const responseData = await response.json().catch(() => null)

      if (!response.ok) {
        const message = responseData?.errors ? JSON.stringify(responseData.errors) : 'Document export failed.'
        throw new Error(message)
      }

      applyPreviewResponse(responseData)
    } catch (exportError) {
      setError(exportError.message)
    } finally {
      setIsFinalizing(false)
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
                  placeholder={DEFAULT_TRANSCRIPT}
                  onChange={(event) => setSourceText(event.target.value)}
                  rows={18}
                  required
                />
                <span className="field-note">
                  This is the only required input. The backend uses it to generate the title, sections, LaTeX, and
                  PDF.
                </span>
              </label>

              <CollapsibleSection title="Advanced options" icon={<SlidersHorizontal size={14} />}>
                <div className="grid-two">
                  <label className="field-group">
                    <span className="field-label">Template</span>
                    <select className="creator-field" value={template} onChange={(e) => setTemplate(e.target.value)}>
                      {TEMPLATE_OPTIONS.map((opt) => (
                        <option key={opt.value} value={opt.value}>
                          {opt.label}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label className="field-group">
                    <span className="field-label">Length</span>
                    <select className="creator-field" value={targetLength} onChange={(e) => setTargetLength(e.target.value)}>
                      {LENGTH_OPTIONS.map((opt) => (
                        <option key={opt.value} value={opt.value}>
                          {opt.label}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label className="field-group">
                    <span className="field-label">Tone</span>
                    <select className="creator-field" value={tone} onChange={(e) => setTone(e.target.value)}>
                      {TONE_OPTIONS.map((opt) => (
                        <option key={opt.value} value={opt.value}>
                          {opt.label}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label className="field-group">
                    <span className="field-label">Max sections (optional)</span>
                    <input
                      className="creator-field"
                      type="number"
                      min={2}
                      max={20}
                      value={maxSections}
                      placeholder="No limit"
                      onChange={(e) => setMaxSections(e.target.value)}
                    />
                  </label>
                </div>
                <label className="options-row">
                  <input
                    type="checkbox"
                    checked={reviewBeforeExport}
                    onChange={(e) => setReviewBeforeExport(e.target.checked)}
                  />
                  <span>Review &amp; edit sections before exporting</span>
                </label>
              </CollapsibleSection>

              <CollapsibleSection title="Additional files" icon={<Upload size={14} />}>
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
              </CollapsibleSection>

              <div className="action-row">
                <button className="primary-button" type="submit" disabled={isGenerating}>
                  {isGenerating ? <LoaderCircle size={16} className="spin" /> : <ArrowDownToLine size={16} />}
                  {isGenerating
                    ? reviewBeforeExport
                      ? 'Generating sections...'
                      : 'Generating PDF...'
                    : reviewBeforeExport
                      ? 'Generate sections to review'
                      : 'Generate preview'}
                </button>
              </div>
            </form>
          </SectionCard>
        </section>

        <section className="preview-column">
          <SectionCard title={viewMode === 'review' ? 'Review sections' : 'PDF preview'} icon={<Layers3 size={18} />}>
            {error ? <div className="alert alert--error">{error}</div> : null}

            {viewMode === 'review' && sectionsDraft ? (
              <>
                <div className="preview-toolbar">
                  <div>
                    <p className="preview-label">Edit before exporting</p>
                    <h3>{documentMeta?.title}</h3>
                  </div>
                  <div className="download-actions">
                    {preview ? (
                      <button type="button" className="secondary-button" onClick={() => setViewMode('preview')}>
                        <Layers3 size={16} /> Back to preview
                      </button>
                    ) : null}
                    <button
                      type="button"
                      className="primary-button"
                      disabled={isFinalizing || sectionsDraft.length === 0}
                      onClick={handleFinalizeExport}
                    >
                      {isFinalizing ? <LoaderCircle size={16} className="spin" /> : <CheckCircle2 size={16} />}
                      {isFinalizing ? 'Exporting...' : 'Finalize & export'}
                    </button>
                  </div>
                </div>

                <div className="section-editor-list">
                  {sectionsDraft.map((section, index) => (
                    <SectionEditorCard
                      key={section._key}
                      section={section}
                      index={index}
                      total={sectionsDraft.length}
                      imageItems={imageItems}
                      isRegenerating={regeneratingIndex === index}
                      onChange={updateSection}
                      onDelete={deleteSection}
                      onMove={moveSection}
                      onRegenerate={regenerateSection}
                    />
                  ))}
                  {sectionsDraft.length === 0 ? (
                    <p className="field-note">All sections were removed — add at least one before exporting.</p>
                  ) : null}
                </div>
              </>
            ) : preview ? (
              <>
                <div className="preview-toolbar">
                  <div>
                    <p className="preview-label">Ready to download</p>
                    <h3>{preview.document.title}</h3>
                  </div>
                  <div className="download-actions">
                    {sectionsDraft ? (
                      <button type="button" className="secondary-button" onClick={() => setViewMode('review')}>
                        <Pencil size={16} /> Edit sections
                      </button>
                    ) : null}
                    <a className="secondary-button" href={preview.pdfUrl} download={preview.filename}>
                      <ArrowDownToLine size={16} /> Download PDF
                    </a>
                    <a className="secondary-button" href={preview.docxUrl} download={preview.docxFilename}>
                      <ArrowDownToLine size={16} /> Download Word
                    </a>
                    <a className="secondary-button" href={preview.texZipUrl} download={preview.texZipFilename}>
                      <ArrowDownToLine size={16} /> Download LaTeX
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
