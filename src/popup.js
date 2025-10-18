import _ from 'lodash'
import axios from 'axios'
import dayjs from 'dayjs'
import * as Excel from 'exceljs/dist/exceljs.min.js'
const supplierBaseUrl = 'https://supplier.coupang.com'
const supplierOrderPathname = '/plan/ticket/reportIssue/CHANGE_PO_INBOUND_DATE_AND_FC'

// DOM 초기화
const $ = (s) => document.querySelector(s)

// 입고예정일 기본값 설정
document.addEventListener('DOMContentLoaded', () => {
  const getOrderBtn = $('#getOrderBtn')
  const getCenterBtn = $('#getCenterBtn')
  const q_date = $('#q_date')
  const q_change_date = $('#q_date_change')
  const q_stores = $('#q_stores')
  const fileInput = $('#fileInput')
  const preferredFCs = $('#preferredFCs')
  const weights = $('#weights')

  if (q_date) {
    q_date.value = dayjs().format('YYYY-MM-DD')
  }

  if (getOrderBtn) {
    getOrderBtn.addEventListener('click', async () => {
      if (!q_date.value) {
        alert('기존 입고예정일을 선택해주세요.')
        return
      }

      if (!q_change_date.value) {
        alert('변경 입고예정일을 선택해주세요.')
        return
      }

      const auth = await sendMsgGetAuth()

      if (!auth) {
        alert('빅셀에 먼저 로그인해주세요.')
        return
      }

      const orders = await sendMsgOrders(auth, q_date.value, q_stores.value)
      const newOrders = transformOrders(orders)

      if (newOrders) {
        await downloadOrdersExcel(newOrders)
      }
    })

    if (getCenterBtn) {
      // await checkCoupangTab()
      getCenterBtn.addEventListener('click', async () => {
        fileInput.click()
      })
    }

    if (fileInput) {
      fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0]

        if (!file) {
          alert('파일을 선택해주세요.')
          return
        }

        const reader = new FileReader()

        reader.onload = async (e) => {
          try {
            console.log('파일 읽기 시작')
            const arrayBuffer = e.target.result
            const workbook = new Excel.Workbook()
            await workbook.xlsx.load(arrayBuffer)

            const sheet = workbook.getWorksheet(1)
            let orders = []

            // 2행부터 데이터 읽기 (1행은 헤더)
            sheet.eachRow((row, rowNumber) => {
              if (rowNumber > 1) {
                // 헤더 제외
                const store_id = row.getCell(1).value
                const order_no = row.getCell(2).value
                const product_name = row.getCell(3).value
                const item_sku_id = row.getCell(4).value
                const order_qty = row.getCell(5).value
                const confirmed_qty = row.getCell(6).value

                if (order_no) {
                  orders.push({
                    order_no,
                    store_id,
                    product_name,
                    item_sku_id,
                    order_qty,
                    confirmed_qty,
                  })
                }
              }
            })

            orders = _.filter(orders, (order) => order.confirmed_qty > 0)

            if (!q_change_date.value) {
              alert('변경 입고예정일이 선택되지 않았습니다.')
              return
            }

            const orderCenters = await sendMsgGetCenters(orders, q_change_date.value)
            console.log('orderCenters : ', orderCenters)
            const preferred = preferredFCs.value.split(',')
            console.log('preferred : ', preferred)

            const result = decideCenterForBatch(orderCenters, preferred, {
              centerFreq: 1,
              groupFreq: 0.5,
              preference: 1,
            })

            console.log('result : ', result)

            const fcList = result.assignments
            await downloadCenterExcel(fcList)
          } catch (error) {
            console.error('파일 처리 오류:', error)
            alert('파일 처리 실패: ' + error.message)
          }
        }

        reader.onerror = (e) => {
          console.error('파일 읽기 오류:', e)
          alert('파일 읽기 실패')
        }

        reader.readAsArrayBuffer(file)
      })
    }
  }
})

// sleep
const sleep = async (ms) => {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

// 로그인 확인
const sendMsgGetAuth = async () => {
  const message = { action: 'getAuth' }
  const res = chrome.runtime.sendMessage(message)
  return res
}

// 발주서 조회
const sendMsgOrders = async (auth, q_date, q_stores = null) => {
  const message = {
    action: 'getOrders',
    auth: auth,
    q_date: dayjs(q_date).format('YYYYMMDD'),
    q_stores: q_stores,
  }
  const res = chrome.runtime.sendMessage(message)
  return res
}

const downloadOrdersExcel = async (orders) => {
  const message = {
    action: 'downloadOrdersExcel',
    orders,
  }
  const res = await chrome.runtime.sendMessage(message)
  downloadExcel(res)
}

const downloadCenterExcel = async (fcList) => {
  const message = {
    action: 'downloadCenterExcel',
    fcList,
  }
  const res = await chrome.runtime.sendMessage(message)
  downloadExcel(res)
}

// 엑셀 다운로드
const downloadExcel = async (res) => {
  if (res?.success && res?.data) {
    // Base64를 Blob으로 변환
    const byteCharacters = atob(res.data)
    const byteNumbers = new Array(byteCharacters.length)
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i)
    }
    const byteArray = new Uint8Array(byteNumbers)
    const blob = new Blob([byteArray], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })

    // 다운로드 링크 생성 및 클릭
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = res.filename
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)

    // alert('엑셀 다운로드 완료!')
  } else {
    alert('엑셀 다운로드 실패: ' + (res.error || '알 수 없는 오류'))
  }
}

// 서플라이어허브 센터 조회
const sendMsgGetCenters = async (orders, q_change_date) => {
  try {
    const message = {
      action: 'getCenters',
      orders,
      q_change_date,
    }

    // 쿠팡 서플라이어 페이지 탭 찾기
    const tabs = await chrome.tabs.query({
      url: '*://supplier.coupang.com/*',
    })

    if (tabs.length === 0) {
      throw new Error('쿠팡 서플라이어 페이지가 열려있지 않습니다.')
    }

    const targetTab = tabs.find((tab) => tab.active) || tabs[0]

    try {
      const res = await chrome.tabs.sendMessage(targetTab.id, message)
      return res
    } catch (error) {
      // Content script가 로드되지 않은 경우 - background script를 통해 처리
      if (error.message.includes('Could not establish connection')) {
        // Background script를 통해 content script 주입 및 메시지 전송 요청
        const response = await chrome.runtime.sendMessage({
          action: 'injectAndSendMessage',
          tabId: targetTab.id,
          message: message,
        })
        return response
      }
      throw error
    }
  } catch (err) {
    throw new Error(err.message)
  }
}
const toNum = (v) => Math.max(0, Number(v) || 0)
const transformOrders = (orders) => {
  const grouped = _.groupBy(orders, 'item_sku_id')
  const result = []

  _.forEach(grouped, (rows, sku) => {
    const initStock = _.get(
      _.maxBy(rows, (r) => toNum(r.hq_stock_qty)),
      'hq_stock_qty',
      0
    )
    const initIncoming = _.get(
      _.maxBy(rows, (r) => toNum(r.hq_stock_pending_qty)),
      'hq_stock_pending_qty',
      0
    )

    let remainStock = toNum(initStock)
    let remainIncoming = toNum(initIncoming)

    // 처리 순서: 주문 일시가 있으면 그걸로, 없으면 입력 순서 유지
    const sorted = _.sortBy(rows, (r) => r.order_datetime ?? r.order_date ?? 0)

    sorted.forEach((row, idx) => {
      const need = toNum(row.order_qty)

      // 1) 본사 재고 우선 소진
      const useStock = Math.min(need, remainStock)
      remainStock -= useStock

      // 2) 입고예정으로 보충 소진
      const stillNeed = need - useStock
      const useIncoming = Math.min(stillNeed, remainIncoming)
      remainIncoming -= useIncoming

      const confirmed = useStock + useIncoming

      result.push({
        ...row,
        confirmed_qty: confirmed,

        allocated_from_stock: useStock,
        allocated_from_incoming: useIncoming,
        remaining_hq_stock_qty_after: remainStock,
        remaining_hq_stock_pending_qty_after: remainIncoming,
      })
    })
  })

  return result
}

const toUpperSafe = (s) => String(s || '').toUpperCase()

const splitCenter = (fcCode) => {
  const m = String(fcCode).match(/^([A-Za-z]+)(\d+)?$/)
  if (!m) return { group: toUpperSafe(fcCode), num: Number.MAX_SAFE_INTEGER }
  return {
    group: toUpperSafe(m[1]),
    num: m[2] ? parseInt(m[2], 10) : Number.MAX_SAFE_INTEGER,
  }
}

// ["GOY","INC","HOB"] → {GOY:3, INC:2, HOB:1}
const buildPreferenceWeights = (preferred = []) => {
  const result = {}
  const len = preferred.length
  preferred.forEach((p, i) => (result[toUpperSafe(p)] = len - i))
  return result
}

const buildFrequencies = (orders) => {
  const centerFreq = new Map()
  const groupFreq = new Map()
  for (const o of orders) {
    const { fcCode } = o
    const centers = Array.isArray(o.center) ? o.center : []
    for (const c of centers) {
      centerFreq.set(c, (centerFreq.get(c) ?? 0) + 1)
      const { group } = splitCenter(fcCode)
      groupFreq.set(group, (groupFreq.get(group) ?? 0) + 1)
    }
  }
  return { centerFreq, groupFreq }
}

const scoreCenter = (fcCode, prefWeights, centerFreq, groupFreq, w) => {
  const { group } = splitCenter(fcCode)
  return (
    (centerFreq.get(fcCode) ?? 0) * w.centerFreq +
    (groupFreq.get(group) ?? 0) * w.groupFreq +
    (prefWeights[group] ?? 0) * w.preference
  )
}

const pickBestCenter = (centers, prefWeights, centerFreq, groupFreq, w) => {
  const scored = centers.map((c) => {
    const { fcCode } = c
    const { group, num } = splitCenter(fcCode)
    return {
      center: fcCode,
      group,
      num,
      score: scoreCenter(fcCode, prefWeights, centerFreq, groupFreq, w),
      gFreq: groupFreq.get(group) ?? 0,
      cFreq: centerFreq.get(fcCode) ?? 0,
    }
  })

  scored.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score
    if (b.gFreq !== a.gFreq) return b.gFreq - a.gFreq
    if (a.num !== b.num) return a.num - b.num
    return a.center.localeCompare(b.center)
  })

  return scored[0].center
}

const decideCenterForBatch = (orders, preferred = [], wOverrides = {}) => {
  const prefWeights = buildPreferenceWeights(preferred)

  const w = {
    centerFreq: wOverrides.centerFreq ?? 1,
    groupFreq: wOverrides.groupFreq ?? 0.5,
    preference: wOverrides.preference ?? 1,
  }

  const { centerFreq, groupFreq } = buildFrequencies(orders)
  const allCenters = Array.from(centerFreq.keys())

  const batchCenter = pickBestCenter(allCenters, prefWeights, centerFreq, groupFreq, w)

  const assignments = orders.map((o) => {
    const centers = Array.isArray(o.center) ? o.center : []
    const fallback = pickBestCenter(centers, prefWeights, centerFreq, groupFreq, w)
    const finalCenter = centers.includes(batchCenter) ? batchCenter : fallback

    return {
      order_no: o.order_no,
      center: finalCenter,
      reason: centers.includes(batchCenter) ? 'batchCenter' : 'fallback',
      candidates: centers,
    }
  })

  return {
    batchCenter,
    centerFreq: Object.fromEntries(centerFreq),
    groupFreq: Object.fromEntries(groupFreq),
    assignments,
  }
}
