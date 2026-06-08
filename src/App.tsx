import React, { useState, useEffect } from 'react'
import { PublicClientApplication } from '@azure/msal-browser'
import type { Configuration, AccountInfo } from '@azure/msal-browser'
import './App.css'


type Tab = 'overview' | 'settings' | 'teams-reader'


export default function App(): React.JSX.Element {
  const [activeTab, setActiveTab] = useState<Tab>('overview')
  
  // States cho Metrics động (Fluctuating)
  const [activeUsers, setActiveUsers] = useState<number>(1420)
  const [cpuLoad, setCpuLoad] = useState<number>(34)
  const [ramUsage, setRamUsage] = useState<number>(58)
  const [teamsMessagesCount, setTeamsMessagesCount] = useState<number>(0)
  
  // Dữ liệu biểu đồ cột mẫu
  const [chartData, setChartData] = useState<number[]>([65, 45, 75, 55, 85, 40, 95])


  // States cho Settings View
  const [profile, setProfile] = useState({ name: 'Nguyễn Văn Quỳnh', email: 'quynh.dev@quynhprivate.com' })
  const [apiKey] = useState('qp_live_6f3e1a8b9c2d5f0e7a4b8c9d')
  const [showApiKey, setShowApiKey] = useState<boolean>(false)
  const [systemToggles, setSystemToggles] = useState({
    autoRefresh: true,
    emailAlerts: true,
    developerMode: false,
    autoReadTeams: true
  })

  // States cho Tích hợp Teams
  const [webhookUrl, setWebhookUrl] = useState<string>('')
  const [isSendingWebhook, setIsSendingWebhook] = useState<boolean>(false)
  const [webhookStatus, setWebhookStatus] = useState<{ type: 'success' | 'error', text: string } | null>(null)

  // States cho Teams TTS Reader
  const [configClientId, setConfigClientId] = useState<string>(() => localStorage.getItem('teams_client_id') || '')
  const [configTenantId, setConfigTenantId] = useState<string>(() => localStorage.getItem('teams_tenant_id') || 'common')
  const [accessToken, setAccessToken] = useState<string | null>(null)
  const [msalInstance, setMsalInstance] = useState<PublicClientApplication | null>(null)
  const [account, setAccount] = useState<AccountInfo | null>(null)
  const [isLoggedIn, setIsLoggedIn] = useState<boolean>(false)
  const [readLogs, setReadLogs] = useState<{ time: string, text: string }[]>([])
  const [speechLang, setSpeechLang] = useState<string>('vi-VN')



  // Hiệu ứng dao động chỉ số thời gian thực (Real-time Simulation)
  useEffect(() => {
    const interval = setInterval(() => {
      // Chỉ tự động cập nhật khi Auto Refresh được bật ở trang cài đặt
      if (!systemToggles.autoRefresh) return

      // Dao động nhẹ số người dùng hoạt động: +/- 10 người
      setActiveUsers((prev) => prev + Math.floor(Math.random() * 21) - 10)
      
      // Dao động tải CPU: 20% -> 80%
      setCpuLoad((prev) => {
        const next = prev + Math.floor(Math.random() * 15) - 7
        return next < 15 ? 15 : next > 85 ? 85 : next
      })

      // Dao động Ram: 50% -> 70%
      setRamUsage((prev) => {
        const next = prev + Math.floor(Math.random() * 5) - 2
        return next < 45 ? 45 : next > 75 ? 75 : next
      })

      // Tự cập nhật dữ liệu biểu đồ ngẫu nhiên
      setChartData((prev) => prev.map(val => {
        const diff = Math.floor(Math.random() * 15) - 7
        const next = val + diff
        return next < 10 ? 10 : next > 100 ? 100 : next
      }))
    }, 4000)

    return () => clearInterval(interval)
  }, [systemToggles.autoRefresh])

  // Lưu thiết lập cài đặt mô phỏng
  const handleSaveProfile = (e: React.FormEvent): void => {
    e.preventDefault()
    alert('Thông tin tài khoản đã được cập nhật trực tuyến!')
  }

  // Hàm xử lý gửi tin nhắn webhook về Teams
  const handleTestWebhook = async (): Promise<void> => {
    if (!webhookUrl) {
      setWebhookStatus({ type: 'error', text: 'Vui lòng nhập URL Webhook trước khi thử nghiệm.' })
      return
    }

    setIsSendingWebhook(true)
    setWebhookStatus(null)

    const payload = {
      "@type": "MessageCard",
      "@context": "http://schema.org/extensions",
      "themeColor": "4f46e5",
      "summary": "Tín hiệu kiểm tra kết nối từ QuynhPrivate Portal",
      "sections": [{
        "activityTitle": "🔔 Kiểm tra kết nối Teams Webhook",
        "activitySubtitle": "Gửi thành công từ Web App Dashboard",
        "facts": [
          { "name": "Người thực hiện", "value": profile.name },
          { "name": "CPU hiện tại", "value": `${cpuLoad}%` },
          { "name": "Thời gian", "value": new Date().toLocaleString('vi-VN') }
        ],
        "markdown": true
      }]
    }

    try {
      // Teams Incoming Webhook thường chặn CORS khi gọi trực tiếp từ trình duyệt của khách hàng.
      // Dùng mode: 'no-cors' để gửi đi mà không bị trình duyệt chặn (phản hồi sẽ ở dạng opaque).
      await fetch(webhookUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload),
        mode: 'no-cors'
      })

      setWebhookStatus({
        type: 'success',
        text: 'Đã phát lệnh gửi! Hãy kiểm tra trực tiếp kênh Microsoft Teams của bạn.'
      })
    } catch (err: any) {
      setWebhookStatus({
        type: 'error',
        text: `Lỗi kết nối: ${err.message || 'Không xác định'}`
      })
    } finally {
      setIsSendingWebhook(false)
    }
  }


  // Cài đặt giọng đọc Text-to-Speech
  const speakText = (text: string): void => {
    if (!('speechSynthesis' in window)) {
      console.warn('Trình duyệt không hỗ trợ speechSynthesis!')
      return
    }

    // Cancel speech đang chạy dở
    window.speechSynthesis.cancel()

    const utterance = new SpeechSynthesisUtterance(text)
    utterance.lang = speechLang
    
    // Tìm voice tiếng Việt hoặc ngôn ngữ phù hợp
    const voices = window.speechSynthesis.getVoices()
    const matchVoice = voices.find(v => v.lang.toLowerCase().includes(speechLang.toLowerCase()))
    if (matchVoice) {
      utterance.voice = matchVoice
    }
    
    window.speechSynthesis.speak(utterance)
  }

  // Đăng nhập Microsoft
  const handleTeamsLogin = async (): Promise<void> => {
    if (!configClientId) {
      alert('Vui lòng cấu hình Client ID trước khi đăng nhập!')
      return
    }

    // Lưu cấu hình vào localStorage để không phải nhập lại
    localStorage.setItem('teams_client_id', configClientId)
    localStorage.setItem('teams_tenant_id', configTenantId)

    const msalConfig: Configuration = {
      auth: {
        clientId: configClientId,
        authority: `https://login.microsoftonline.com/${configTenantId || 'common'}`,
        redirectUri: window.location.origin
      },
      cache: {
        cacheLocation: 'localStorage'
      }
    }

    try {
      const pca = new PublicClientApplication(msalConfig)
      await pca.initialize()
      
      const loginRequest = {
        scopes: ['User.Read', 'Chat.Read']
      }
      
      const response = await pca.loginPopup(loginRequest)
      setAccessToken(response.accessToken)
      setMsalInstance(pca)
      setAccount(response.account)
      setIsLoggedIn(true)
      
      const loginMsg = `Kết nối thành công tài khoản: ${response.account.username}`
      setReadLogs(prev => [{ time: new Date().toLocaleTimeString(), text: loginMsg }, ...prev])
      speakText('Đăng nhập Microsoft Teams thành công!')
    } catch (err: any) {
      console.error(err)
      alert(`Đăng nhập thất bại: ${err.message || err.toString()}`)
    }
  }

  // Đăng xuất
  const handleTeamsLogout = (): void => {
    setAccessToken(null)
    setAccount(null)
    setIsLoggedIn(false)
    if (msalInstance && account) {
      msalInstance.logoutPopup({
        account: account,
        postLogoutRedirectUri: window.location.origin
      }).catch(err => console.error('Logout error:', err))
    }
    setReadLogs(prev => [{ time: new Date().toLocaleTimeString(), text: 'Đã đăng xuất tài khoản.' }, ...prev])
    speakText('Đã đăng xuất tài khoản!')
  }

  // Tiến trình chạy ngầm quét tin nhắn mới (Graph API Polling)
  useEffect(() => {
    if (!isLoggedIn || !accessToken || !systemToggles.autoReadTeams) return

    const lastSeenMessageIds: { [chatId: string]: string } = {}
    let isFirstRun = true
    let activeInterval: any = null


    const checkNewMessages = async (): Promise<void> => {
      try {
        const chatsResponse = await fetch('https://graph.microsoft.com/v1.0/me/chats?$top=20', {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        })

        if (!chatsResponse.ok) {
          if (chatsResponse.status === 401) {
            handleTeamsLogout()
            return
          }
          throw new Error(`Chats status ${chatsResponse.status}`)
        }

        const chatsData = await chatsResponse.json()
        const chats = chatsData.value || []

        for (const chat of chats) {
          // Chỉ lấy tin nhắn 1-on-1 cá nhân
          if (chat.chatType !== 'oneOnOne') continue

          const chatId = chat.id

          const messagesResponse = await fetch(`https://graph.microsoft.com/v1.0/chats/${chatId}/messages?$top=5`, {
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Content-Type': 'application/json'
            }
          })

          if (!messagesResponse.ok) continue

          const messagesData = await messagesResponse.json()
          const messages = messagesData.value || []
          
          if (messages.length === 0) continue

          // Sắp xếp tăng dần theo thời gian tạo
          const sortedMessages = [...messages].sort((a, b) => 
            new Date(a.createdDateTime).getTime() - new Date(b.createdDateTime).getTime()
          )

          const latestMessage = sortedMessages[sortedMessages.length - 1]
          const latestMessageId = latestMessage.id
          
          // Lần đầu chạy, lưu lại tin nhắn cuối cùng để làm mốc, tránh đọc hàng loạt tin nhắn cũ
          if (!lastSeenMessageIds[chatId]) {
            lastSeenMessageIds[chatId] = latestMessageId
            continue
          }

          if (lastSeenMessageIds[chatId] !== latestMessageId) {
            const lastSeenIdx = sortedMessages.findIndex(m => m.id === lastSeenMessageIds[chatId])
            const newMessages = lastSeenIdx === -1 
              ? [latestMessage]
              : sortedMessages.slice(lastSeenIdx + 1)

            lastSeenMessageIds[chatId] = latestMessageId

            if (isFirstRun) continue

            for (const msg of newMessages) {
              // Lọc các tin nhắn dạng văn bản và do người khác gửi (không phải chính mình gửi)
              if (msg.messageType !== 'message' || !msg.body || !msg.body.content) continue
              if (msg.from && msg.from.user && msg.from.user.id === account?.localAccountId) continue

              // Loại bỏ các thẻ HTML trong nội dung tin nhắn Teams
              const tempDiv = document.createElement('div')
              tempDiv.innerHTML = msg.body.content
              const plainText = tempDiv.textContent || tempDiv.innerText || ''

              if (!plainText.trim()) continue

              const senderName = msg.from?.user?.displayName || 'Đồng nghiệp'
              const textToSpeak = `Tin nhắn từ ${senderName}: ${plainText}`

              // Ghi log vào màn hình
              const timeString = new Date(msg.createdDateTime).toLocaleTimeString()
              setReadLogs(prev => [
                { time: timeString, text: `Đọc to: "${textToSpeak}"` },
                ...prev.slice(0, 49)
              ])

              // Đọc tin nhắn thành tiếng
              speakText(textToSpeak)

              // Tăng số lượng tin nhắn đã đọc
              setTeamsMessagesCount(prev => prev + 1)
            }
          }
        }
        
        isFirstRun = false
      } catch (err: any) {
        console.error('Lỗi khi kiểm tra tin nhắn Teams:', err)
      }
    }

    // Chạy ngay lần đầu và thiết lập lặp lại mỗi 5 giây
    checkNewMessages()
    activeInterval = setInterval(checkNewMessages, 5000)

    return () => {
      if (activeInterval) clearInterval(activeInterval)
    }
  }, [isLoggedIn, accessToken, systemToggles.autoReadTeams])


  return (

    <div className="web-app-container">
      {/* Sidebar Navigation */}
      <aside className="web-sidebar">
        <div className="sidebar-header">
          <span className="logo-icon">💠</span>
          <span className="brand-name">QuynhPrivate Portal</span>
        </div>

        <ul className="sidebar-menu">
          <li 
            className={`menu-item ${activeTab === 'overview' ? 'active' : ''}`}
            onClick={(): void => setActiveTab('overview')}
          >
            📊 Tổng quan Dashboard
          </li>
          <li 
            className={`menu-item ${activeTab === 'settings' ? 'active' : ''}`}
            onClick={(): void => setActiveTab('settings')}
          >
            ⚙️ Cấu hình Hệ thống
          </li>
          <li 
            className={`menu-item ${activeTab === 'teams-reader' ? 'active' : ''}`}
            onClick={(): void => setActiveTab('teams-reader')}
          >
            📢 Trình Đọc Teams (TTS)
          </li>
        </ul>


        <div className="sidebar-footer">
          <div className="user-badge">
            <div className="avatar">Q</div>
            <div className="user-info">
              <span className="username">{profile.name}</span>
              <span className="user-role">Administrator</span>
            </div>
          </div>
        </div>
      </aside>

      {/* Main Panel */}
      <main className="web-main-content">
        {/* Header */}
        <header className="web-header">
          <div className="header-left">
            <h2>
              {activeTab === 'overview' && 'Bảng Điều Khiển Tổng Quan'}
              {activeTab === 'settings' && 'Thiết Lập Hệ Thống & Tài Khoản'}
              {activeTab === 'teams-reader' && 'Trình Đọc Tin Nhắn Cá Nhân Teams (TTS)'}
            </h2>

          </div>
          <div className="header-right">
            <div className="status-indicator">
              <span className="status-dot"></span>
              {systemToggles.autoRefresh ? 'Thời gian thực (Live)' : 'Đã tạm dừng Live'}
            </div>
          </div>
        </header>

        {/* Content Body */}
        <div className="web-content-body">
          <div className="view-container">
            
            {/* TAB 1: OVERVIEW */}
            {activeTab === 'overview' && (
              <div className="overview-tab">
                {/* Metrics Card Grid */}
                <div className="metrics-grid">
                  <div className="metric-card">
                    <div className="metric-header">
                      <span className="metric-title">Người dùng trực tuyến</span>
                      <span className="metric-icon">👥</span>
                    </div>
                    <p className="metric-value">{activeUsers.toLocaleString()}</p>
                    <span className="metric-change positive">▲ +4.2% so với giờ trước</span>
                  </div>

                  <div className="metric-card">
                    <div className="metric-header">
                      <span className="metric-title">Tải lượng CPU</span>
                      <span className="metric-icon">💻</span>
                    </div>
                    <p className="metric-value">{cpuLoad}%</p>
                    <span className={`metric-change ${cpuLoad > 70 ? 'negative' : 'positive'}`}>
                      {cpuLoad > 70 ? '▲ Cảnh báo quá tải' : '▼ CPU đang ở mức an toàn'}
                    </span>
                  </div>

                  <div className="metric-card">
                    <div className="metric-header">
                      <span className="metric-title">Bộ nhớ RAM đã dùng</span>
                      <span className="metric-icon">💾</span>
                    </div>
                    <p className="metric-value">{ramUsage}%</p>
                    <span className="metric-change positive">● Tổng dung lượng 16 GB</span>
                  </div>

                  <div className="metric-card">
                    <div className="metric-header">
                      <span className="metric-title">Tin nhắn Teams đã đọc</span>
                      <span className="metric-icon">💬</span>
                    </div>
                    <p className="metric-value">{teamsMessagesCount}</p>
                    <span className="metric-change positive">
                      ● Đang tự động lắng nghe
                    </span>
                  </div>
                </div>

                {/* Charts Section */}
                <div className="charts-section">
                  <div className="chart-card">
                    <h3 className="chart-title">Thống kê Truy cập Tuần qua (Simulated)</h3>
                    <div className="mock-chart-container">
                      {chartData.map((val, idx) => {
                        const labels = ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN']
                        return (
                          <div className="chart-bar-wrapper" key={idx}>
                            <div 
                              className="chart-bar" 
                              style={{ height: `${val}%` }}
                            >
                              <div className="chart-tooltip">{val}k requests</div>
                            </div>
                            <span className="chart-label">{labels[idx]}</span>
                          </div>
                        )
                      })}
                    </div>
                  </div>

                  <div className="chart-card">
                    <h3 className="chart-title">Tình Trạng Hệ Thống</h3>
                    <div className="stats-list">
                      <div className="stat-item">
                        <div className="stat-info">
                          <span className="stat-label">Thời gian hoạt động (Uptime)</span>
                          <span className="stat-value">99.98%</span>
                        </div>
                        <div className="stat-progress-bar">
                          <div className="stat-progress-fill success" style={{ width: '99.98%' }}></div>
                        </div>
                      </div>

                      <div className="stat-item">
                        <div className="stat-info">
                          <span className="stat-label">Dung lượng ổ đĩa SSD</span>
                          <span className="stat-value">62%</span>
                        </div>
                        <div className="stat-progress-bar">
                          <div className="stat-progress-fill warning" style={{ width: '62%' }}></div>
                        </div>
                      </div>

                      <div className="stat-item">
                        <div className="stat-info">
                          <span className="stat-label">Băng thông sử dụng</span>
                          <span className="stat-value">28%</span>
                        </div>
                        <div className="stat-progress-bar">
                          <div className="stat-progress-fill" style={{ width: '28%' }}></div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>


                {/* MS Teams Integration Row (Adjusted Dashboard) */}
                <div style={{ marginTop: '20px' }}>
                  <div className="metric-card" style={{ margin: 0, justifyContent: 'space-between', minHeight: '180px' }}>
                    <div className="metric-header">
                      <span className="metric-title">Trạng thái đọc Teams (TTS)</span>
                      <span className={`status-badge ${isLoggedIn ? 'active' : 'inactive'}`} style={{ margin: 0 }}>
                        {isLoggedIn ? '● Đang lắng nghe' : '✕ Chưa kết nối'}
                      </span>
                    </div>
                    
                    {isLoggedIn ? (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', flexGrow: 1, marginTop: '8px' }}>
                        <div style={{ fontSize: '13px', color: 'var(--text-muted)' }}>
                          Tài khoản: <strong style={{ color: 'var(--text-main)' }}>{account?.name || account?.username}</strong>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', flexGrow: 1, marginTop: '6px' }}>
                          <span style={{ fontSize: '12px', fontWeight: 'bold', color: 'var(--text-muted)' }}>Nhật ký gần đây:</span>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            {readLogs.slice(0, 2).map((log, idx) => (
                              <div key={idx} style={{ display: 'flex', gap: '8px', borderBottom: '1px solid rgba(255,255,255,0.03)', paddingBottom: '4px', fontSize: '12px' }}>
                                <span style={{ color: '#60a5fa', fontFamily: 'monospace' }}>[{log.time}]</span>
                                <span style={{ textOverflow: 'ellipsis', overflow: 'hidden', whiteSpace: 'nowrap', color: 'var(--text-main)' }}>{log.text}</span>
                              </div>
                            ))}
                            {readLogs.length === 0 && (
                              <span style={{ color: 'var(--text-muted)', fontStyle: 'italic', fontSize: '12px' }}>Chưa có tin nhắn cá nhân nào được đọc.</span>
                            )}
                          </div>
                        </div>
                      </div>
                    ) : (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', justifyContent: 'center', alignItems: 'center', flexGrow: 1, padding: '10px 0' }}>
                        <p style={{ fontSize: '13px', color: 'var(--text-muted)', textAlign: 'center', margin: 0 }}>
                          Vui lòng thiết lập và đăng nhập tài khoản Microsoft Teams để bắt đầu tự động đọc tin nhắn.
                        </p>
                        <button 
                          onClick={(): void => setActiveTab('teams-reader')}
                          className="btn-submit"
                          style={{ alignSelf: 'center', padding: '6px 12px', fontSize: '12px', background: '#5b5fc7', marginTop: '8px' }}
                        >
                          Cấu hình kết nối
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}




            {/* TAB 3: SETTINGS */}
            {activeTab === 'settings' && (
              <div className="settings-tab">
                <div className="settings-layout">
                  {/* Account Information Card */}
                  <div className="settings-card">
                    <h3>Thông tin tài khoản</h3>
                    <form className="settings-form" onSubmit={handleSaveProfile}>
                      <div className="form-group">
                        <label>Họ và tên</label>
                        <input 
                          type="text" 
                          value={profile.name} 
                          onChange={(e): void => setProfile({ ...profile, name: e.target.value })} 
                          required
                        />
                      </div>
                      
                      <div className="form-group">
                        <label>Email liên lạc</label>
                        <input 
                          type="email" 
                          value={profile.email} 
                          onChange={(e): void => setProfile({ ...profile, email: e.target.value })} 
                          required
                        />
                      </div>

                      <div className="form-group">
                        <label>Khóa API cổng dịch vụ (API Key)</label>
                        <div className="api-key-input-wrapper">
                          <input 
                            type={showApiKey ? 'text' : 'password'} 
                            value={apiKey} 
                            disabled 
                          />
                          <button 
                            type="button" 
                            className="btn-toggle-visibility"
                            onClick={(): void => setShowApiKey(!showApiKey)}
                          >
                            {showApiKey ? 'Ẩn' : 'Hiện'}
                          </button>
                        </div>
                      </div>

                      <button type="submit" className="btn-submit">Lưu thông tin</button>
                    </form>
                  </div>

                  {/* System Settings Card */}
                  <div className="settings-card">
                    <h3>Cấu hình hệ thống Dashboard</h3>
                    <div className="settings-form">
                      <div className="switch-group">
                        <div className="switch-label-wrapper">
                          <span className="switch-title">Cập nhật dữ liệu thời gian thực</span>
                          <span className="switch-description">Tự động quét và cập nhật số liệu trực tuyến sau mỗi 4 giây.</span>
                        </div>
                        <label className="switch">
                          <input 
                            type="checkbox" 
                            checked={systemToggles.autoRefresh}
                            onChange={(e): void => setSystemToggles({ ...systemToggles, autoRefresh: e.target.checked })}
                          />
                          <span className="slider"></span>
                        </label>
                      </div>

                      <div className="switch-group">
                        <div className="switch-label-wrapper">
                          <span className="switch-title">Thông báo qua Email</span>
                          <span className="switch-description">Gửi cảnh báo bảo mật về tài khoản khi CPU hoặc RAM vượt ngưỡng 90%.</span>
                        </div>
                        <label className="switch">
                          <input 
                            type="checkbox" 
                            checked={systemToggles.emailAlerts}
                            onChange={(e): void => setSystemToggles({ ...systemToggles, emailAlerts: e.target.checked })}
                          />
                          <span className="slider"></span>
                        </label>
                      </div>

                      <div className="switch-group">
                        <div className="switch-label-wrapper">
                          <span className="switch-title">Chế độ nhà phát triển (Developer Mode)</span>
                          <span className="switch-description">Bật logs chuyên sâu ở Console trình duyệt và các API phụ.</span>
                        </div>
                        <label className="switch">
                          <input 
                            type="checkbox" 
                            checked={systemToggles.developerMode}
                            onChange={(e): void => setSystemToggles({ ...systemToggles, developerMode: e.target.checked })}
                          />
                          <span className="slider"></span>
                        </label>
                      </div>

                      <div className="switch-group">
                        <div className="switch-label-wrapper">
                          <span className="switch-title">Tự động đọc tin nhắn Teams (TTS)</span>
                          <span className="switch-description">Sử dụng Text-to-Speech đọc to các tin nhắn 1-on-1 Teams mới.</span>
                        </div>
                        <label className="switch">
                          <input 
                            type="checkbox" 
                            checked={systemToggles.autoReadTeams}
                            onChange={(e): void => setSystemToggles({ ...systemToggles, autoReadTeams: e.target.checked })}
                          />
                          <span className="slider"></span>
                        </label>
                      </div>


                      <button 
                        type="button" 
                        className="btn-submit"
                        onClick={(): void => alert('Cấu hình hệ thống đã được cập nhật thành công!')}
                      >
                        Áp dụng cấu hình
                      </button>
                    </div>
                  </div>
                </div>

                {/* Teams Webhook Card */}
                <div className="settings-card" style={{ marginTop: '24px' }}>
                  <h3>Tích hợp Microsoft Teams (Webhook)</h3>
                  <div className="settings-form">
                    <div className="form-group">
                      <label>Địa chỉ Teams Incoming Webhook URL</label>
                      <input 
                        type="text" 
                        placeholder="https://your-tenant.webhook.office.com/webhookb2/..." 
                        value={webhookUrl}
                        onChange={(e): void => setWebhookUrl(e.target.value)}
                      />
                      <p className="hint">
                        Nhập URL Webhook từ cấu hình Incoming Webhook (hoặc Workflows app) của kênh Teams để đẩy thông báo trực tiếp.
                      </p>
                    </div>
                    <button 
                      type="button" 
                      className="btn-submit"
                      onClick={handleTestWebhook}
                      disabled={isSendingWebhook}
                    >
                      {isSendingWebhook ? 'Đang gửi...' : 'Gửi báo cáo Test về Teams'}
                    </button>
                    {webhookStatus && (
                      <div className={`webhook-status-text ${webhookStatus.type}`}>
                        {webhookStatus.text}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}


            {/* TAB 4: TEAMS READER (TTS) */}

            {activeTab === 'teams-reader' && (

              <div className="teams-reader-tab">
                <div className="settings-layout">
                  {/* Cấu hình Đăng nhập */}
                  <div className="settings-card">
                    <h3>Xác thực Microsoft Teams</h3>
                    {!isLoggedIn ? (
                      <div className="settings-form">
                        <div className="form-group">
                          <label>Microsoft Entra Application (Client) ID</label>
                          <input 
                            type="text" 
                            placeholder="Nhập Client ID của ứng dụng đăng ký trên Azure AD" 
                            value={configClientId}
                            onChange={(e): void => setConfigClientId(e.target.value)}
                          />
                        </div>
                        <div className="form-group">
                          <label>Tenant ID</label>
                          <input 
                            type="text" 
                            placeholder="Mặc định: common (Đa tenant)" 
                            value={configTenantId}
                            onChange={(e): void => setConfigTenantId(e.target.value)}
                          />
                        </div>
                        <div className="form-group">
                          <label>Địa chỉ Redirect URI</label>
                          <input 
                            type="text" 
                            value={window.location.origin} 
                            disabled 
                          />
                          <p className="hint">
                            Vui lòng cấu hình Web Redirect URI này trong trang cấu hình Authentication của ứng dụng trên Azure Portal.
                          </p>
                        </div>
                        <button 
                          type="button" 
                          className="btn-submit"
                          style={{ backgroundColor: '#5b5fc7' }}
                          onClick={handleTeamsLogin}
                        >
                          🔑 Đăng nhập Teams
                        </button>
                      </div>
                    ) : (
                      <div className="settings-form">
                        <div className="user-profile-display" style={{ display: 'flex', alignItems: 'center', gap: '16px', padding: '16px', background: 'rgba(255,255,255,0.03)', borderRadius: '8px', border: '1px solid var(--color-border)' }}>
                          <div className="avatar" style={{ width: '48px', height: '48px', fontSize: '18px', background: 'linear-gradient(135deg, #7c3aed, #5b5fc7)', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', borderRadius: '50%', fontWeight: 'bold' }}>
                            {account?.name ? account.name.charAt(0) : 'U'}
                          </div>
                          <div style={{ display: 'flex', flexDirection: 'column' }}>
                            <span style={{ fontWeight: 'bold', fontSize: '15px' }}>{account?.name || 'Tài khoản Microsoft'}</span>
                            <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>{account?.username}</span>
                          </div>
                        </div>

                        <div className="switch-group" style={{ marginTop: '16px' }}>
                          <div className="switch-label-wrapper">
                            <span className="switch-title">Tự động đọc tin nhắn (TTS)</span>
                            <span className="switch-description">Khi có tin nhắn cá nhân (1-on-1) mới, hệ thống tự động phát giọng đọc.</span>
                          </div>
                          <label className="switch">
                            <input 
                              type="checkbox" 
                              checked={systemToggles.autoReadTeams}
                              onChange={(e): void => setSystemToggles({ ...systemToggles, autoReadTeams: e.target.checked })}
                            />
                            <span className="slider"></span>
                          </label>
                        </div>

                        <div className="form-group">
                          <label>Ngôn ngữ & Giọng đọc</label>
                          <select 
                            value={speechLang}
                            onChange={(e): void => setSpeechLang(e.target.value)}
                          >
                            <option value="vi-VN">Tiếng Việt (vi-VN)</option>
                            <option value="en-US">Tiếng Anh (en-US)</option>
                            <option value="ja-JP">Tiếng Nhật (ja-JP)</option>
                          </select>
                        </div>

                        <div style={{ display: 'flex', gap: '12px', marginTop: '12px' }}>
                          <button 
                            type="button" 
                            className="btn-submit"
                            style={{ backgroundColor: '#5b5fc7' }}
                            onClick={(): void => speakText('Tính năng chuyển văn bản thành giọng nói đang hoạt động tốt!')}
                          >
                            🔊 Thử giọng đọc
                          </button>
                          <button 
                            type="button" 
                            className="btn-submit"
                            style={{ backgroundColor: 'transparent', border: '1px solid var(--color-error)', color: 'var(--color-error)' }}
                            onClick={handleTeamsLogout}
                          >
                            Đăng xuất
                          </button>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* Nhật ký đọc tin nhắn */}
                  <div className="settings-card" style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
                    <h3>Nhật ký Đọc Tin Nhắn (TTS Logs) - Tổng số: {teamsMessagesCount}</h3>
                    <div 
                      className="tts-logs-container" 
                      style={{ 
                        flexGrow: 1, 
                        background: 'rgba(0,0,0,0.2)', 
                        borderRadius: '8px', 
                        padding: '16px', 
                        maxHeight: '340px', 
                        overflowY: 'auto',
                        border: '1px solid var(--color-border)'
                      }}
                    >
                      {readLogs.length > 0 ? (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                          {readLogs.map((log, idx) => (
                            <div 
                              key={idx} 
                              style={{ 
                                display: 'flex', 
                                gap: '10px', 
                                fontSize: '13px', 
                                borderBottom: '1px solid rgba(255,255,255,0.03)', 
                                paddingBottom: '8px' 
                              }}
                            >
                              <span style={{ color: '#60a5fa', fontFamily: 'monospace' }}>[{log.time}]</span>
                              <span style={{ color: 'var(--text-main)' }}>{log.text}</span>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <p style={{ color: 'var(--text-muted)', fontSize: '13px', textAlign: 'center', marginTop: '40px' }}>
                          {!isLoggedIn ? 'Vui lòng đăng nhập để bắt đầu lắng nghe tin nhắn.' : 'Chưa nhận được tin nhắn cá nhân mới nào.'}
                        </p>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>

        </div>
      </main>
    </div>
  )
}
