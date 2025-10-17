import _ from 'lodash'
import { Buffer } from 'buffer'
import pako from 'pako'
import backend from './api/backend.local'
import axios from 'axios'
import * as Excel from 'exceljs/dist/exceljs.min.js'

chrome.runtime.onInstalled.addListener(() => {
  console.log('Extension installed!')
})

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  switch (message.action) {
    case 'getAuth':
      getAuth()
        .then((auth) => {
          sendResponse(auth)
        })
        .catch((err) => {
          sendResponse({ error: err.message })
        })

      return true

    case 'getOrders':
      getCoupangRocketOrders({ auth: message.auth, q_date: message.q_date })
        .then((res) => {
          sendResponse(res)
        })
        .catch((err) => {
          sendResponse({ error: err.message })
        })

      return true

    case 'getSupplierCenter':
      getSupplierCenter({ orders: message.orders, q_date: message.q_date })
        .then((res) => {
          sendResponse(res)
        })
        .catch((err) => {
          sendResponse({ error: err.message })
        })
      return true

    case 'downloadOrdersExcel':
      downloadOrdersExcel({ orders: message.orders })
        .then((res) => {
          sendResponse(res)
        })
        .catch((err) => {
          sendResponse({ error: err.message })
        })
      return true

    case 'downloadCenterExcel':
      downloadCenterExcel({ orders: message.orders })
        .then((res) => {
          sendResponse(res)
        })
        .catch((err) => {
          sendResponse({ error: err.message })
        })
      return true

    case 'searchAndClickOrders':
      searchAndClickOrders({ tabId: message.tabId, orders: message.orders })
        .then((res) => {
          sendResponse(res)
        })
        .catch((err) => {
          sendResponse({ error: err.message })
        })
      return true

    default:
      sendResponse({ error: '잘못된 요청입니다.' })
      return true
  }
})

const getAuth = async () => {
  try {
    const res = await axios
      .get(`${backend.loginCheck}`, {
        headers: {
          'Content-Type': 'application/json',
        },
      })
      .then(async (res) => {
        if (res?.status === 200) {
          const data = await res.data
          if (data.status === 'succeed') {
            const decodedData = atob(data.body)
            const buffer = Buffer.from(decodedData, 'binary')
            const decompressData = pako.inflate(buffer, { to: 'string' })
            return decompressData
          } else {
            return null
          }
        } else {
          return null
        }
      })
      .then((res) => {
        return JSON.parse(res)
      })

    if (!res?.id_token) {
      throw new Error('빅셀에 먼저 로그인해주세요.')
    }

    return res
  } catch (error) {
    // Sentry.captureException(error)
    throw new Error(error.message)
  }
}

const getCoupangRocketOrders = async ({ auth, q_date, q_stores = '' }) => {
  try {
    const { id_token } = auth

    const params = {
      q_date_tpyes: 'exp_recv_date',
      q_sale_date_from: q_date,
      q_sale_date_to: q_date,
      q_stores: q_stores,
    }

    const q_params = new URLSearchParams(params).toString()

    const orders = await axios
      .post(
        `${backend.getCoupangRocketOrders}?${q_params}`,
        {},
        {
          headers: {
            Authorization: id_token,
          },
        }
      )
      .then(async (res) => {
        if (res?.status === 200) {
          const data = await res.data
          if (data.status === 'succeed') {
            const decodedData = atob(data.body)
            const buffer = Buffer.from(decodedData, 'binary')
            const decompressData = pako.inflate(buffer, { to: 'string' })
            return decompressData
          } else {
            return null
          }
        } else {
          return null
        }
      })
      .then((res) => {
        return JSON.parse(res).orders
      })

    return orders
  } catch (error) {
    // Sentry.captureException(error)
    throw new Error(error.message)
  }
}

const getSupplierCenter = async ({ orders, q_date }) => {
  try {
    const url = `https://supplier.coupang.com/plan/v1/ticket/poEditing/getNewCalendarAvailableList`

    const params = {
      poId: orders[1].order_no,
      partnerId: orders[1].store_id,
      skuId: '',
      reason: '',
      request_model: '',
      datetime: new Date().getTime(),
    }

    const { data } = await axios.get(url, {
      params,
      withCredentials: true,
    })

    const { body: centers } = data

    const filteredCenters = centers.filter((c) => c.status === 'AVAILABLE' && c.edd === q_date)
    console.log('filteredCenters : ', filteredCenters)

    return filteredCenters
  } catch (error) {
    // Sentry.captureException(error)
    throw new Error(error.message)
  }
}

const downloadOrdersExcel = async ({ orders }) => {
  try {
    const addHeaders = [
      '발주번호',
      '상품명',
      'SKU ID',
      '발주수량',
      '확정수량',
      '남은본사재고',
      '본사재고',
      '입고예정수량',
      '입고예정일',
      '입고센터',
    ]
    const addHeaderWidths = [20, 50, 20, 20, 20, 20, 20, 20, 20]

    // ExcelJS로 워크북 생성
    const wb = new Excel.Workbook()
    const sheet = wb.addWorksheet('로켓서플라이어 주문')

    // 헤더 추가
    const headerRow = sheet.addRow(addHeaders)
    headerRow.eachCell((cell, colNum) => {
      sheet.getColumn(colNum).width = addHeaderWidths[colNum - 1]
      cell.font = { bold: true }
      cell.alignment = { vertical: 'middle', horizontal: 'center' }
    })

    orders.forEach((order) => {
      let {
        order_no,
        item_sku_name,
        item_sku_id,
        order_qty,
        confirmed_qty,
        remaining_hq_stock_qty_after,
        hq_stock_qty,
        hq_stock_pending_qty,
        pending_detail,
        logis_center_name,
      } = order

      if (pending_detail) {
        pending_detail = _.chain()
          .get(JSON.parse(pending_detail), 'pending_detail')
          .map((item) => `${item?.hq_stock_pending_arrived_date} - ${item?.hq_stock_pending_qty}`)
          .join('\r\n')
          .value()
      }

      const rowData = [
        order_no, // 발주번호
        item_sku_name, // 상품명
        item_sku_id, // SKU ID
        order_qty, // 발주수량
        confirmed_qty, // 확정수량
        remaining_hq_stock_qty_after, // 남은본사재고
        hq_stock_qty, // 본사재고
        hq_stock_pending_qty, // 입고예정수량
        pending_detail, // 입고예정일
        logis_center_name, // 변경전 납품센터
      ]
      sheet.addRow(rowData)
    })

    // 엑셀 파일 생성
    const buffer = await wb.xlsx.writeBuffer()

    // ArrayBuffer를 Base64로 변환하여 popup으로 반환
    const base64 = btoa(
      new Uint8Array(buffer).reduce((data, byte) => data + String.fromCharCode(byte), '')
    )

    return {
      success: true,
      message: '엑셀 생성 완료',
      data: base64,
      filename: `로켓서플라이어_주문_${new Date().getTime()}.xlsx`,
    }
  } catch (error) {
    throw new Error(error.message)
  }
}

const downloadCenterExcel = async ({ orders }) => {
  try {
    const addHeaders = [
      '발주번호',
      '상품명',
      'SKU ID',
      '발주수량',
      '확정수량',
      '남은본사재고',
      '본사재고',
      '입고예정수량',
      '입고예정일',
      '입고센터',
      '변경센터',
      '후보센터',
    ]
    const addHeaderWidths = [20, 50, 20, 20, 20, 20, 20, 20, 20, 20, 30, 50]

    // public 폴더의 엑셀 템플릿 파일 경로 가져오기
    const templateUrl = chrome.runtime.getURL('supplier_rocket_order.xlsx')

    // 템플릿 파일 읽어오기
    const response = await fetch(templateUrl)
    const arrayBuffer = await response.arrayBuffer()

    // ExcelJS로 워크북 로드
    const wb = new Excel.Workbook()
    await wb.xlsx.load(arrayBuffer)

    const sheet = wb.getWorksheet(1)

    // F열부터 추가 헤더 입력 (2행)
    const headerRow = sheet.getRow(2)
    addHeaders.forEach((header, index) => {
      const colIndex = 6 + index // F열부터 시작
      const cell = headerRow.getCell(colIndex)
      cell.value = header
      cell.font = { bold: true }
      cell.alignment = { vertical: 'middle', horizontal: 'center' }
      sheet.getColumn(colIndex).width = addHeaderWidths[index]
    })

    // 3행부터 데이터 입력
    orders.forEach((order, index) => {
      const row = sheet.getRow(3 + index)
      row.getCell(6).value = order.order_no || ''
      row.getCell(7).value = order.item_sku_name || ''
      row.getCell(8).value = order.item_sku_id || ''
      row.getCell(9).value = order.order_qty || 0
      row.getCell(10).value = order.confirmed_qty || 0
      row.getCell(11).value = order.remaining_hq_stock_qty_after || 0
      row.getCell(12).value = order.hq_stock_qty || 0
      row.getCell(13).value = order.hq_stock_pending_qty || 0
      row.getCell(14).value = order.hq_stock_pending_arrived_date || ''
      row.getCell(15).value = order.logis_center_name || ''
      row.getCell(16).value = order.changed_center || '' // 변경센터
      row.getCell(17).value = order.candidate_centers || '' // 후보센터
    })

    // 엑셀 파일 생성
    const buffer = await wb.xlsx.writeBuffer()

    // ArrayBuffer를 Base64로 변환하여 popup으로 반환
    const base64 = btoa(
      new Uint8Array(buffer).reduce((data, byte) => data + String.fromCharCode(byte), '')
    )

    return {
      success: true,
      message: '엑셀 생성 완료',
      data: base64,
      filename: `로켓서플라이어_센터변경_${new Date().getTime()}.xlsx`,
    }
  } catch (error) {
    throw new Error(error.message)
  }
}

const searchAndClickOrders = async ({ tabId, orders }) => {
  try {
    console.log('background - searchAndClickOrders 호출:', { tabId, orders })

    // 탭 정보 확인
    const tab = await chrome.tabs.get(tabId)
    console.log('background - 탭 정보:', tab.url)

    // content script로 메시지 전송 시도
    try {
      const response = await chrome.tabs.sendMessage(tabId, {
        action: 'searchAndClick',
        orders: orders,
      })

      console.log('background - content script 응답:', response)
      return response
    } catch (error) {
      // content script가 로드되지 않은 경우 - 동적 inject 시도
      if (error.message.includes('Could not establish connection')) {
        console.log('background - content script inject 시도...')

        try {
          // content script 동적 inject
          await chrome.scripting.executeScript({
            target: { tabId: tabId },
            files: ['content.js'],
          })

          console.log('background - inject 성공, 재시도...')

          // 잠시 대기 후 재시도
          await new Promise((resolve) => setTimeout(resolve, 1000))

          const response = await chrome.tabs.sendMessage(tabId, {
            action: 'searchAndClick',
            orders: orders,
          })

          console.log('background - 재시도 성공:', response)
          return response
        } catch (injectError) {
          console.error('background - inject 실패:', injectError)
          throw new Error('Content script 로드 실패. 페이지를 새로고침한 후 다시 시도해주세요.')
        }
      }
      throw error
    }
  } catch (error) {
    console.error('background - searchAndClickOrders 오류:', error)
    throw error
  }
}
