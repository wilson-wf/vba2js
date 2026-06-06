import { createApp, ref, onMounted } from 'vue'
import axios from 'axios'

createApp({
  setup() {
    const fileInput = ref(null)
    const selectedFile = ref(null)
    const isDragover = ref(false)
    const useLLM = ref(false)
    const isConverting = ref(false)
    const progress = ref(0)
    const currentStage = ref('')
    const convertSuccess = ref(false)
    const convertError = ref('')
    const downloadUrl = ref('')
    const fileId = ref('')

    const llmConfig = ref({
      apiKey: '',
      endpoint: 'https://api.openai.com/v1/chat/completions',
      model: 'gpt-4'
    })
    
    const llmStatus = ref('')
    const llmStatusClass = ref('')

    const triggerUpload = () => {
      fileInput.value?.click()
    }

    const handleFileSelect = (event) => {
      const file = event.target.files[0]
      if (file) {
        selectedFile.value = file
        convertSuccess.value = false
        convertError.value = ''
      }
    }

    const handleDrop = (event) => {
      isDragover.value = false
      const file = event.dataTransfer.files[0]
      if (file && file.name.endsWith('.xlsm')) {
        selectedFile.value = file
        convertSuccess.value = false
        convertError.value = ''
      }
    }

    const uploadFile = async (file) => {
      const formData = new FormData()
      formData.append('file', file)
      
      const response = await axios.post('/api/upload', formData, {
        headers: { 'Content-Type': 'multipart/form-data' }
      })
      
      return response.data.file_id
    }

    const startConvert = async () => {
      if (!selectedFile.value) return

      isConverting.value = true
      progress.value = 0
      currentStage.value = '上传文件...'
      convertSuccess.value = false
      convertError.value = ''

      try {
        currentStage.value = '上传文件中...'
        fileId.value = await uploadFile(selectedFile.value)
        
        currentStage.value = '开始转换...'
        
        await axios.post(`/api/convert/${fileId.value}`, {
          use_llm: useLLM.value
        })

        await pollStatus()

        downloadUrl.value = `/api/download/${fileId.value}`
        convertSuccess.value = true
        currentStage.value = '转换完成'
        progress.value = 100
      } catch (error) {
        convertError.value = error.response?.data?.error || '转换失败，请重试'
        currentStage.value = '转换失败'
      } finally {
        isConverting.value = false
      }
    }

    const pollStatus = async () => {
      return new Promise((resolve) => {
        const interval = setInterval(async () => {
          try {
            const response = await axios.get(`/api/convert/status/${fileId.value}`)
            const status = response.data
            
            progress.value = status.progress
            currentStage.value = getStageText(status.stage)

            if (status.status === 'completed') {
              clearInterval(interval)
              resolve()
            } else if (status.status === 'failed') {
              clearInterval(interval)
              throw new Error(status.error_message)
            }
          } catch (error) {
            clearInterval(interval)
            throw error
          }
        }, 500)
      })
    }

    const getStageText = (stage) => {
      const texts = {
        'pending': '等待中',
        'parsing': '解析VBA代码...',
        'converting': '转换语法...',
        'testing': '测试验证...',
        'generating': '生成文件...',
        'completed': '完成',
        'failed': '失败'
      }
      return texts[stage] || stage
    }

    const saveLLMConfig = async () => {
      try {
        await axios.post('/api/config/llm', {
          api_key: llmConfig.value.apiKey,
          endpoint: llmConfig.value.endpoint,
          model: llmConfig.value.model
        })
        llmStatus.value = '配置保存成功'
        llmStatusClass.value = 'status-success'
        setTimeout(() => { llmStatus.value = '' }, 3000)
      } catch (error) {
        llmStatus.value = '保存失败'
        llmStatusClass.value = 'status-error'
      }
    }

    const loadLLMConfig = async () => {
      try {
        const response = await axios.get('/api/config/llm')
        llmConfig.value.endpoint = response.data.endpoint
        llmConfig.value.model = response.data.model
        useLLM.value = response.data.has_api_key
      } catch (error) {
        console.log('Failed to load LLM config')
      }
    }

    onMounted(() => {
      loadLLMConfig()
    })

    return {
      fileInput,
      selectedFile,
      isDragover,
      useLLM,
      isConverting,
      progress,
      currentStage,
      convertSuccess,
      convertError,
      downloadUrl,
      llmConfig,
      llmStatus,
      llmStatusClass,
      triggerUpload,
      handleFileSelect,
      handleDrop,
      startConvert,
      saveLLMConfig
    }
  }
}).mount('body')