import { useEffect, useRef, useState } from 'react'
import { BsChatDotsFill, BsSend, BsX } from 'react-icons/bs'
import { AzureOpenAI } from 'openai'

type ChatMessage = {
  id: number
  role: 'assistant' | 'user'
  text: string
}

const initialMessages: ChatMessage[] = [
  { id: 1, role: 'assistant', text: 'Hello. I can help with translation. Give your message in the text area below with the target language specified.' },
]

const azureOpenAiEndpoint = import.meta.env.VITE_AZURE_OPENAI_ENDPOINT?.replace(/\/$/, '')
const azureOpenAiApiKey = import.meta.env.VITE_AZURE_OPENAI_API_KEY
const azureOpenAiDeployment = import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT
const azureOpenAiApiVersion = import.meta.env.VITE_AZURE_OPENAI_API_VERSION || '2024-10-21'
const targetLanguage = import.meta.env.VITE_AZURE_OPENAI_TARGET_LANGUAGE || 'English'

async function translateText(userPrompt: string): Promise<string> {
  if (!azureOpenAiEndpoint || !azureOpenAiApiKey || !azureOpenAiDeployment) {
    throw new Error('Azure OpenAI is not configured. Please contact the Administrator.')
  }

  const client = new AzureOpenAI({
    apiKey: azureOpenAiApiKey,
    apiVersion: azureOpenAiApiVersion,
    deployment: azureOpenAiDeployment,
    endpoint: azureOpenAiEndpoint,
    dangerouslyAllowBrowser: true
  })

  const messages = [
    {
      role: 'system' as const,
      content: `Use the user's complete prompt as the source for the translation. If the prompt specifies a target language, translate into that language; otherwise, translate into ${targetLanguage}. Return only the translated text, without explanations or quotation marks.`
    },
    {
      role: 'user' as const,
      content: userPrompt
    }
  ];

  const response = await client.chat.completions.create({
    messages,
    model: azureOpenAiDeployment,
    temperature: 0.3,
  })

  const translation = response.choices[0]?.message?.content?.trim()
  if (!translation) {
    throw new Error('Azure OpenAI returned an empty translation.')
  }

  return translation
}

export function SupportChatbot() {
  const [isOpen, setIsOpen] = useState(false)
  const [draft, setDraft] = useState('')
  const [messages, setMessages] = useState<ChatMessage[]>(initialMessages)
  const [isTyping, setIsTyping] = useState(false)
  const nextMessageId = useRef(2)
  const messagesEndRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages, isTyping])

  const sendMessage = async () => {
    const text = draft.trim()
    if (!text || isTyping) return

    setMessages((current) => [...current, { id: nextMessageId.current++, role: 'user', text }])
    setDraft('')
    setIsTyping(true)

    try {
      const translation = await translateText(text)
      setMessages((current) => [...current, {
        id: nextMessageId.current++,
        role: 'assistant',
        text: translation,
      }])
    } catch (error) {
      setMessages((current) => [...current, {
        id: nextMessageId.current++,
        role: 'assistant',
        text: error instanceof Error ? error.message : 'Unable to translate the message right now.',
      }])
    } finally {
      setIsTyping(false)
    }
  }

  const handleKeyDown = (event: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault()
      sendMessage()
    }
  }

  return <div className="support-chatbot">
    {isOpen && <section className="chat-panel" role="dialog" aria-modal="false" aria-labelledby="support-chat-title">
      <header className="chat-panel-header">
        <div className="chat-panel-brand">
          <span className="chat-panel-icon" aria-hidden="true"><BsChatDotsFill /></span>
          <div>
            <h2 id="support-chat-title">Translate Assistant</h2>
            <span><i className="chat-status-dot" /> Ready to help</span>
          </div>
        </div>
        <button className="chat-close-button" type="button" aria-label="Close chat" onClick={() => setIsOpen(false)}><BsX /></button>
      </header>

      <div className="chat-messages" aria-live="polite">
        <div className="chat-date-divider"><span>Today</span></div>
        {messages.map((message) => <div className={`chat-message-row ${message.role}`} key={message.id}>
          <div className="chat-message-bubble">{message.text}</div>
        </div>)}
        {isTyping && <div className="chat-message-row assistant"><div className="chat-message-bubble chat-typing" aria-label="Assistant is typing"><span /><span /><span /></div></div>}
        <div ref={messagesEndRef} />
      </div>

      <div className="chat-composer">
        <textarea value={draft} onChange={(event) => setDraft(event.target.value)} onKeyDown={handleKeyDown} placeholder="Write a message..." aria-label="Message" rows={1} />
        <button type="button" className="chat-send-button" aria-label="Send message" disabled={!draft.trim() || isTyping}><BsSend /></button>
      </div>
      <p className="chat-disclaimer">Assistant responses may need verification.</p>
    </section>}

    <button className={`chat-launcher ${isOpen ? 'is-open' : ''}`} type="button" aria-label={isOpen ? 'Close assistant' : 'Open assistant'} aria-expanded={isOpen} onClick={() => setIsOpen((current) => !current)}>
      {isOpen ? <BsX /> : <BsChatDotsFill />}
      {!isOpen && <span className="chat-launcher-ping" aria-hidden="true" />}
    </button>
  </div>
}
