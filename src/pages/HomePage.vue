<template>
  <div class="p-6 relative">
    <!-- Template Dropdown -->
    <div
      class="bg-gray-100 rounded-lg shadow-sm p-4 mb-4 flex items-center justify-center"
    >
      <el-icon class="text-blue-600 mr-3 text-xl flex-shrink-0"
        ><Document
      /></el-icon>
      <el-dropdown>
        <span
          class="flex items-center text-blue-700 font-medium cursor-pointer border-b border-gray-200 pb-2"
        >
          Create a draft using a template
          <el-icon class="ml-2 pb-3"><ArrowDown /></el-icon>
        </span>
        <template #dropdown>
          <el-dropdown-menu>
            <el-dropdown-item>Service Agreement</el-dropdown-item>
            <el-dropdown-item>Non-Disclosure Agreement</el-dropdown-item>
            <el-dropdown-item>Employment Contract</el-dropdown-item>
          </el-dropdown-menu>
        </template>
      </el-dropdown>
    </div>

    <!-- Review Options Cards -->
    <div class="space-y-3">
      <!-- Partial Review Card -->
      <div
        class="bg-blue-50 rounded-lg p-4 cursor-pointer hover:bg-blue-100 transition-colors"
      >
        <div class="flex items-center">
          <div class="mr-3 flex-shrink-0 flex items-center justify-center w-6">
            <el-icon class="text-blue-600 text-lg"><Search /></el-icon>
          </div>
          <div>
            <h3 class="font-medium text-blue-800 flex">Partial Review</h3>
            <p class="text-sm text-blue-700">
              Document summary and analysis of toxic provision
            </p>
          </div>
        </div>
      </div>

      <!-- Full Review Card - Selected -->
      <div class="bg-blue-600 rounded-lg p-4 cursor-pointer">
        <div class="flex items-center">
          <div class="mr-3 flex-shrink-0 flex items-center justify-center w-6">
            <el-icon class="text-white text-lg"><Document /></el-icon>
          </div>
          <div>
            <h3 class="font-medium text-white flex">Full Review</h3>
            <p class="text-sm text-blue-100">
              Analyze the toxicity of the selected text
            </p>
          </div>
        </div>
      </div>

      <!-- Recommend Next Content -->
      <div
        class="bg-blue-50 rounded-lg p-4 cursor-pointer hover:bg-blue-100 transition-colors"
      >
        <div class="flex items-center">
          <div class="mr-3 flex-shrink-0 flex items-center justify-center w-6">
            <el-icon class="text-blue-600 text-lg"><ArrowRight /></el-icon>
          </div>
          <div>
            <h3 class="font-medium text-blue-800 flex">
              Recommend Next Content
            </h3>
            <p class="text-sm text-blue-700">
              Suggestion for next provision based on context
            </p>
          </div>
        </div>
      </div>

      <!-- Full Document Analysis Result -->
      <div class="rounded-lg p-4 cursor-pointer">
        <div class="flex items-center justify-between">
          <div class="flex items-center">
            <div
              class="mr-3 flex-shrink-0 flex items-center justify-center w-6"
            >
              <el-icon class="text-blue-600 text-lg"><Document /></el-icon>
            </div>
            <h3 class="font-medium text-gray-800 flex">
              Full document analysis result
            </h3>
          </div>
          <div class="flex items-center">
            <div
              class="mr-3 flex-shrink-0 flex items-center justify-center w-6 gap-1"
            >
              <el-icon class="!text-red-600 text-lg"><InfoFilled /></el-icon>
              <p class="text-sm text-red-700 border-b border-red-700">Report</p>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Document Summary -->
    <div class="mt-6 bg-blue-50 rounded-lg p-4">
      <h3 class="font-bold text-gray-800 mb-2">Document Summary</h3>
      <p class="text-sm text-blue-800 mb-2">
        This contract is a service provision contract between Sample Enterprise
        Co., Ltd. (Party A) and Test Company Co., Ltd. (Party B), and includes
        the following main content:
      </p>
    </div>

    <!-- Toxic Provisions Section (new) -->
    <div class="mt-4 bg-red-50 rounded-lg p-4">
      <div class="flex items-center mb-2">
        <el-icon class="text-red-600 mr-2 text-lg flex-shrink-0"
          ><Warning
        /></el-icon>
        <h3 class="font-bold text-gray-800">Discovery of Toxic Provisions</h3>
      </div>
      <ul class="text-sm text-red-800 space-y-1 list-none">
        <li>
          - Article 2, Paragraph 1: Unilateral Liability Electricity Provisions
        </li>
        <li>
          - Article 2, Paragraph 2: Imposition of Liability for Excessive
          Damages
        </li>
        <li>
          - Article 2, Paragraph 3: Right to Unilaterally Terminate a Contract
        </li>
        <li>- Article 2, Paragraph 4: Unfair Cost Reimbursement Provisions</li>
      </ul>
      <p class="text-sm text-red-800 mt-2">
        To edit a toxic clause, select the section and use the Review Section
        function
      </p>
    </div>

    <!-- Disclaimer -->
    <div class="mt-2 text-center">
      <p class="text-xs text-red-600">
        AI does not provide legal advice and users retain final review
        responsibility.
      </p>
    </div>

    <!-- Action Buttons -->
    <div class="mt-4 mb-6 w-full flex justify-between fixed bottom-0">
      <el-button>Cancel</el-button>
      <el-button class="mr-12 !bg-blue-600" type="primary"
        >Apply Changes</el-button
      >
    </div>
  </div>
</template>

<script lang="ts" setup>
import {
  ArrowDown,
  ArrowRight,
  Document,
  InfoFilled,
  Search,
  Warning
} from '@element-plus/icons-vue'

import { ElMessage } from 'element-plus'
import { onBeforeMount, onMounted, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import { useRouter } from 'vue-router'

import API from '@/api'

import { promptDbInstance } from '@/store/promtStore'
import { buildInPrompt } from '@/utils/constant'

import { checkAuth } from '@/utils/common'
import { localStorageKey } from '@/utils/enum'
import useSettingForm from '@/utils/settingForm'
import { settingPreset } from '@/utils/settingPreset'

const { t } = useI18n()

const { settingForm } = useSettingForm()

// system prompt
const systemPrompt = ref('')
const systemPromptSelected = ref('')
const systemPromptList = ref<IStringKeyMap[]>([])
const addSystemPromptVisible = ref(false)
const addSystemPromptAlias = ref('')
const addSystemPromptValue = ref('')
const removeSystemPromptVisible = ref(false)
const removeSystemPromptValue = ref<any[]>([])

// user prompt
const prompt = ref('')
const promptSelected = ref('')
const promptList = ref<IStringKeyMap[]>([])
const addPromptVisible = ref(false)
const addPromptAlias = ref('')
const addPromptValue = ref('')
const removePromptVisible = ref(false)
const removePromptValue = ref<any[]>([])

// result
const result = ref('res')
const loading = ref(false)
const router = useRouter()
const historyDialog = ref<any[]>([])

const jsonIssue = ref(false)
const errorIssue = ref(false)

// insert type
const insertType = ref<insertTypes>('replace')
// const insertTypeList = ['replace', 'newLine', 'NoAction'].map(item => ({
//   label: t(item),
//   value: item
// }))

// Add this to your setup script section where other refs are defined
const contractType = ref('general')
const contractTypeList = [
  { label: 'General', value: 'general' },
  { label: 'Service Agreement', value: 'service' },
  { label: 'Employment', value: 'employment' },
  { label: 'Non-Disclosure', value: 'nda' },
  { label: 'Sale of Goods', value: 'sale' },
  { label: 'Lease', value: 'lease' },
  { label: 'Software License', value: 'software' },
  { label: 'Consulting', value: 'consulting' }
]

async function getSystemPromptList() {
  const table = promptDbInstance.table('systemPrompt')
  const list = (await table.toArray()) as unknown as IStringKeyMap[]
  systemPromptList.value = list
}

async function addSystemPrompt() {
  const table = promptDbInstance.table('systemPrompt')
  await table.put({
    key: addSystemPromptAlias.value,
    value: addSystemPromptValue.value
  })
  addSystemPromptVisible.value = false
  getSystemPromptList()
}

async function removeSystemPrompt() {
  removeSystemPromptVisible.value = false
  const table = promptDbInstance.table('systemPrompt')
  for (const value of removeSystemPromptValue.value) {
    await table.delete(value)
  }
  removeSystemPromptValue.value = []
  getSystemPromptList()
}

async function removePrompt() {
  removePromptVisible.value = false
  const table = promptDbInstance.table('userPrompt')
  for (const value of removePromptValue.value) {
    await table.delete(value)
  }
  removePromptValue.value = []
  getPromptList()
}

function handelSystemPromptChange(val: string) {
  systemPrompt.value = val
  localStorage.setItem(localStorageKey.defaultSystemPrompt, val)
}

async function getPromptList() {
  const table = promptDbInstance.table('userPrompt')
  const list = (await table.toArray()) as unknown as IStringKeyMap[]
  promptList.value = list
}

async function addPrompt() {
  const table = promptDbInstance.table('userPrompt')
  await table.put({
    key: addPromptAlias.value,
    value: addPromptValue.value
  })
  addPromptVisible.value = false
  getPromptList()
}

function handelPromptChange(val: string) {
  prompt.value = val
  localStorage.setItem(localStorageKey.defaultPrompt, val)
}

const addWatch = () => {
  watch(
    () => settingForm.value.replyLanguage,
    () => {
      localStorage.setItem('replyLanguage', settingForm.value.replyLanguage)
    }
  )
}

async function initData() {
  insertType.value =
    (localStorage.getItem(localStorageKey.insertType) as insertTypes) ||
    'replace'
  systemPrompt.value =
    'Act as AI Legal Assistance. You are helping to review contract and recommend contract terms'
  await getSystemPromptList()
  if (systemPromptList.value.find(item => item.value === systemPrompt.value)) {
    systemPromptSelected.value = systemPrompt.value
  }
  prompt.value = localStorage.getItem(localStorageKey.defaultPrompt) || ''
  await getPromptList()
  contractType.value =
    localStorage.getItem(localStorageKey.contractType) || 'general'
  if (promptList.value.find(item => item.value === prompt.value)) {
    promptSelected.value = prompt.value
  }
}

function handleContractTypeChange(val: string) {
  contractType.value = val
  localStorage.setItem(localStorageKey.contractType, val)
}

function handelInsertTypeChange(val: insertTypes) {
  insertType.value = val
  localStorage.setItem(localStorageKey.insertType, val)
  // Trigger text selection handler to update prompt when insert type changes
  handleTextSelection()
}

async function template(taskType: keyof typeof buildInPrompt | 'custom') {
  loading.value = true
  let systemMessage
  let userMessage = ''
  const getSeletedText = async () => {
    return Word.run(async context => {
      const range = context.document.getSelection()
      range.load('text')
      await context.sync()
      return range.text
    })
  }
  const getDocumentText = async () => {
    return Word.run(async context => {
      const body = context.document.body // Get the document body
      body.load('text') // Load text content
      await context.sync()
      return body.text // Return the full document text
    })
  }

  const previousText = await getPreviousText()

  const documentText = await getDocumentText()

  const selectedText = await getSeletedText()

  console.log(selectedText)
  // console.log(previousText, 'previousText')
  if (taskType === 'custom') {
    if (systemPrompt.value.includes('{language}')) {
      systemMessage = systemPrompt.value.replace(
        '{language}',
        settingForm.value.replyLanguage
      )
    } else {
      systemMessage = systemPrompt.value
    }

    // Append documentText when selectedText is available
    if (userMessage.includes('{text}')) {
      userMessage = userMessage.replace('{text}', selectedText)
    } else if (selectedText) {
      if (insertType.value === 'replace') {
        userMessage = `Reply in ${settingForm.value.replyLanguage}
    This is a ${contractType.value} contract. Based on the full context, suggest a suitable replacement for the following section:

      **Full Contract:**
      "${documentText}"

      **Section to Replace:**
      "${selectedText}"

    Ensure the replacement aligns with the contract's style and legal accuracy for ${contractType.value} contracts.
    Provide an improved version of the clause with enhanced legal coverage. Only return the new improved version without explanation.`
      } else if (insertType.value === 'append') {
        userMessage = `Reply in ${settingForm.value.replyLanguage}
    This is a ${contractType.value} contract. Based on the full context, suggest suitable additional text to append to the following section:

      **Full Contract:**
      "${documentText}"

      **Section to Append To:**
      "${selectedText}"

    Ensure the appended text flows naturally from the existing text and aligns with the contract's style and legal accuracy for ${contractType.value} contracts.
    Only return the text to be appended without explanation.`
      } else if (insertType.value === 'newLine') {
        userMessage = `Reply in ${settingForm.value.replyLanguage}
    This is a ${contractType.value} contract. Based on the full context, suggest a new clause or paragraph to be inserted after the following section:

      **Full Contract:**
      "${documentText}"

      **Section After Which to Insert:**
      "${selectedText}"

    Ensure the new clause aligns with the contract's style and legal accuracy for ${contractType.value} contracts.
    Only return the new clause without explanation.`
      } else {
        userMessage = `Reply in ${settingForm.value.replyLanguage}
    This is a ${contractType.value} contract. Based on the full context, suggest a suitable ${insertType.value} for the following section:

      **Full Contract:**
      "${documentText}"

      **Section to Modify:**
      "${selectedText}"

    Ensure the update aligns with the contract's style and legal accuracy for ${contractType.value} contracts.
    Provide an improved version of the clause with enhanced legal coverage. Only return the new improved version without explanation.`
      }
    } else if (insertType.value === 'newLine') {
      userMessage = `Reply in ${settingForm.value.replyLanguage}

        [Clause Category]:
        ${documentText}

        [Section After Which to Insert]:
        ${previousText}

        [Insert Type]:
        ${insertType.value}

        [New Clause]:
        ${prompt.value}

        Provide an improved version of the clause with enhanced legal coverage. Only return the new improved version without explanation.`
    } else {
      userMessage = `Reply in ${settingForm.value.replyLanguage} ${prompt.value}`
    }
  } else {
    systemMessage = buildInPrompt[taskType].system(
      settingForm.value.replyLanguage
    )
    userMessage = buildInPrompt[taskType].user(
      selectedText,
      settingForm.value.replyLanguage
    )
  }
  if (
    settingForm.value.api === 'official' &&
    settingForm.value.officialAPIKey
  ) {
    console.log(userMessage)
    const config = API.official.setConfig(
      settingForm.value.officialAPIKey,
      settingForm.value.officialBasePath
    )
    historyDialog.value = [
      {
        role: 'system',
        content: systemMessage
      },
      {
        role: 'user',
        content: userMessage
      }
    ]
    await API.official.createChatCompletionStream({
      config,
      messages: historyDialog.value,
      result,
      historyDialog,
      errorIssue,
      loading,
      maxTokens: settingForm.value.officialMaxTokens,
      temperature: settingForm.value.officialTemperature,
      model:
        settingForm.value.officialCustomModel ||
        settingForm.value.officialModelSelect
    })
  } else if (settingForm.value.api === 'groq' && settingForm.value.groqAPIKey) {
    historyDialog.value = [
      {
        role: 'system',
        content: systemMessage
      },
      {
        role: 'user',
        content: userMessage
      }
    ]
    await API.groq.createChatCompletionStream({
      groqAPIKey: settingForm.value.groqAPIKey,
      groqModel:
        settingForm.value.groqCustomModel || settingForm.value.groqModelSelect,
      messages: historyDialog.value,
      result,
      historyDialog,
      errorIssue,
      loading,
      maxTokens: settingForm.value.officialMaxTokens,
      temperature: settingForm.value.officialTemperature
    })
  } else if (
    settingForm.value.api === 'azure' &&
    settingForm.value.azureAPIKey
  ) {
    historyDialog.value = [
      {
        role: 'system',
        content: systemMessage
      },
      {
        role: 'user',
        content: userMessage
      }
    ]
    await API.azure.createChatCompletionStream({
      azureAPIKey: settingForm.value.azureAPIKey,
      azureAPIEndpoint: settingForm.value.azureAPIEndpoint,
      azureDeploymentName: settingForm.value.azureDeploymentName,
      messages: historyDialog.value,
      result,
      historyDialog,
      errorIssue,
      loading,
      maxTokens: settingForm.value.azureMaxTokens,
      temperature: settingForm.value.azureTemperature
    })
  } else if (
    settingForm.value.api === 'gemini' &&
    settingForm.value.geminiAPIKey
  ) {
    historyDialog.value = [
      {
        role: 'user',
        parts: [
          {
            text: systemMessage + '\n' + userMessage
          }
        ]
      },
      {
        role: 'model',
        parts: [
          {
            text: 'Hi, what can I help you?'
          }
        ]
      }
    ]
    await API.gemini.createChatCompletionStream({
      geminiAPIKey: settingForm.value.geminiAPIKey,
      messages: userMessage,
      result,
      historyDialog,
      errorIssue,
      loading,
      maxTokens: settingForm.value.geminiMaxTokens,
      temperature: settingForm.value.geminiTemperature,
      geminiModel:
        settingForm.value.geminiCustomModel ||
        settingForm.value.geminiModelSelect
    })
  } else if (
    settingForm.value.api === 'ollama' &&
    settingForm.value.ollamaEndpoint
  ) {
    historyDialog.value = [
      {
        role: 'user',
        content: systemMessage + '\n' + userMessage
      }
    ]
    await API.ollama.createChatCompletionStream({
      ollamaEndpoint: settingForm.value.ollamaEndpoint,
      ollamaModel:
        settingForm.value.ollamaCustomModel ||
        settingForm.value.ollamaModelSelect,
      messages: historyDialog.value,
      result,
      historyDialog,
      errorIssue,
      loading,
      temperature: settingForm.value.ollamaTemperature
    })
  } else {
    ElMessage.error('Set API Key or Access Token first')
    return
  }
  if (errorIssue.value === true) {
    errorIssue.value = false
    ElMessage.error('Something is wrong')
    return
  }
  if (!jsonIssue.value) {
    API.common.insertResult(result, insertType)
  }
}

function checkApiKey() {
  const auth = {
    type: settingForm.value.api,
    apiKey: settingForm.value.officialAPIKey,
    azureAPIKey: settingForm.value.azureAPIKey,
    geminiAPIKey: settingForm.value.geminiAPIKey,
    groqAPIKey: settingForm.value.groqAPIKey
  }
  if (!checkAuth(auth)) {
    ElMessage.error('Set API Key or Access Token first')
    return false
  }
  return true
}

const actionList = Object.keys(buildInPrompt) as (keyof typeof buildInPrompt)[]

async function performAction(action: keyof typeof buildInPrompt) {
  if (!checkApiKey()) return
  template(action)
}

function settings() {
  router.push('/settings')
}

function StartChat() {
  if (!checkApiKey()) return
  template('custom')
}

async function continueChat() {
  if (!checkApiKey()) return
  loading.value = true
  try {
    switch (settingForm.value.api) {
      case 'official':
        historyDialog.value.push({
          role: 'user',
          content: 'continue'
        })

        await API.official.createChatCompletionStream({
          config: API.official.setConfig(
            settingForm.value.officialAPIKey,
            settingForm.value.officialBasePath
          ),
          messages: historyDialog.value,
          result,
          historyDialog,
          errorIssue,
          loading,
          maxTokens: settingForm.value.officialMaxTokens,
          temperature: settingForm.value.officialTemperature,
          model:
            settingForm.value.officialCustomModel ||
            settingForm.value.officialModelSelect
        })
        break
      case 'groq':
        historyDialog.value.push({
          role: 'user',
          content: 'continue'
        })
        await API.groq.createChatCompletionStream({
          groqAPIKey: settingForm.value.groqAPIKey,
          groqModel:
            settingForm.value.groqCustomModel ||
            settingForm.value.groqModelSelect,
          messages: historyDialog.value,
          result,
          historyDialog,
          errorIssue,
          loading,
          maxTokens: settingForm.value.officialMaxTokens,
          temperature: settingForm.value.officialTemperature
        })
        break
      case 'azure':
        historyDialog.value.push({
          role: 'user',
          content: 'continue'
        })
        await API.azure.createChatCompletionStream({
          azureAPIKey: settingForm.value.azureAPIKey,
          azureAPIEndpoint: settingForm.value.azureAPIEndpoint,
          azureDeploymentName: settingForm.value.azureDeploymentName,
          messages: historyDialog.value,
          result,
          historyDialog,
          errorIssue,
          loading,
          maxTokens: settingForm.value.azureMaxTokens,
          temperature: settingForm.value.azureTemperature
        })
        break
      case 'gemini':
        historyDialog.value.push(
          ...[
            {
              role: 'user',
              parts: [
                {
                  text: 'continue'
                }
              ]
            },
            {
              role: 'model',
              parts: [
                {
                  text: 'OK, I will continue to help you.'
                }
              ]
            }
          ]
        )
        await API.gemini.createChatCompletionStream({
          geminiAPIKey: settingForm.value.geminiAPIKey,
          messages: 'continue',
          result,
          historyDialog,
          errorIssue,
          loading,
          maxTokens: settingForm.value.geminiMaxTokens,
          temperature: settingForm.value.geminiTemperature,
          geminiModel:
            settingForm.value.geminiCustomModel ||
            settingForm.value.geminiModelSelect
        })
        break
      case 'ollama':
        historyDialog.value.push({
          role: 'user',
          content: 'continue'
        })
        await API.ollama.createChatCompletionStream({
          ollamaEndpoint: settingForm.value.ollamaEndpoint,
          ollamaModel:
            settingForm.value.ollamaCustomModel ||
            settingForm.value.ollamaModelSelect,
          messages: historyDialog.value,
          result,
          historyDialog,
          errorIssue,
          loading,
          temperature: settingForm.value.ollamaTemperature
        })
    }
  } catch (error) {
    result.value = String(error)
    errorIssue.value = true
    console.error(error)
  }
  if (errorIssue.value === true) {
    errorIssue.value = false
    ElMessage.error('Something is wrong')
    return
  }
  API.common.insertResult(result, insertType)
}

onBeforeMount(() => {
  addWatch()
  initData()
})

const detectLanguage = () => {
  const userLang = navigator.language.split('-')[0]
  const availableLanguages = settingPreset.replyLanguage.optionList!.map(
    item => item.value
  )

  // Extract primary language (e.g., "en" from "en-US")
  const primaryLang = userLang.split('-')[0]

  // Set the language if it's available, else default to English
  settingForm.value.replyLanguage = availableLanguages.includes(primaryLang)
    ? primaryLang
    : settingPreset.replyLanguage.defaultValue
}

// ✅ Detect selected text and update `settingForm.value.replyLanguage`const handleTextSelection = async () => {

const handleTextSelection = async () => {
  return Word.run(async context => {
    const range = context.document.getSelection()
    range.load('text')
    await context.sync()

    // Get current insert type from local storage
    const currentInsertType = localStorage.getItem('insertType') || 'replace'
    // Different prompt handling based on insert type
    if (
      currentInsertType === 'replace' &&
      range.text &&
      range.text.trim() !== ''
    ) {
      // For replace: use the selected text
      prompt.value = `I would like a new term to replace this term: ${range.text}. Please return the new term only.`
    } else if (currentInsertType === 'newLine') {
      // For newLine: keep whatever is in the prompt field
      // Do nothing to preserve user's existing prompt
      // const previousText = await getPreviousText()
      // const text = previousText ?? range.text
      prompt.value = `
        [Role]:
        You are an expert contract lawyer specializing in drafting legally sound and enforceable contract clauses. Your task is to generate a new clause based on the given category, ensuring it aligns with industry best practices and legal standards.

        [Contract Type]:
        "${contractType.value}"

        [Key Considerations]:
        - Clearly define rights and obligations for both parties.
        - Ensure legal enforceability and risk mitigation.
        - Align with industry standards and best practices.
        - Tailor to applicable jurisdiction (if detected).

        [Preferred Output]:`
    } else {
      // When no text is selected and not in append mode
      prompt.value = ''
    }
  })
}

const getPreviousText = async () => {
  try {
    return await Word.run(async context => {
      // Get the current selection
      const selection = context.document.getSelection()

      // Get the current paragraph containing the selection
      const currentParagraph = selection.paragraphs.getFirst()
      currentParagraph.load('text')

      // Get all paragraphs in the document
      const paragraphs = context.document.body.paragraphs
      paragraphs.load('items')

      await context.sync()

      // Find the index of the current paragraph
      let currentIndex = -1
      for (let i = 0; i < paragraphs.items.length; i++) {
        if (paragraphs.items[i].text === currentParagraph.text) {
          currentIndex = i
          break
        }
      }

      // Get the previous paragraph text if it exists
      if (currentIndex > 0) {
        const previousParagraph = paragraphs.items[currentIndex - 1]
        return previousParagraph.text.trim()
      } else {
        return '' // No previous paragraph found
      }
    })
  } catch (error) {
    console.error('Error getting previous text:', error)
    return ''
  }
}

const setupWordTracking = () => {
  Office.onReady()
    .then(() => {
      if (Office.context.document) {
        Office.context.document.addHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          async () => {
            await handleTextSelection() // Auto-update `prompt.value`
          },
          result => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              console.error('Failed to attach event handler:', result.error)
            }
          }
        )
      } else {
        console.error('Office.context.document is not available.')
      }
    })
    .catch(err => {
      console.error('Office.js failed to initialize:', err)
    })
}

onMounted(() => {
  detectLanguage()
  setupWordTracking()
})
</script>

<style scoped>
.container {
  display: flex;
  flex-direction: column;
  align-items: center;
  margin-top: 50px;
}

.input-group {
  display: flex;
  align-items: center;
  margin-bottom: 20px;
}

.api-button {
  margin-left: 10px;
  border-radius: 10px;
}

.result-group {
  width: 100%;
}
</style>
