import _ from 'lodash'
import { Buffer } from 'buffer'
import pako from 'pako'
import backend from './api/backend.prod'
import axios from 'axios'
import * as Excel from 'exceljs/dist/exceljs.min.js'
import dayjs from 'dayjs'

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

    case 'getProductsStocks':
      getCoupangProductStocks({
        auth: message.auth,
        q_change_date: message.q_change_date,
        sku_list: message.sku_list,
      })
        .then((res) => {
          sendResponse(res)
        })
        .catch((err) => {
          sendResponse({ error: err.message })
        })
      return true

    // case 'getOrders':
    //   getCoupangRocketOrders({ auth: message.auth, q_date: message.q_date })
    //     .then((res) => {
    //       sendResponse(res)
    //     })
    //     .catch((err) => {
    //       sendResponse({ error: err.message })
    //     })

    //   return true

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
      downloadCenterExcel({ fcList: message.fcList })
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

    case 'injectAndSendMessage':
      injectAndSendMessage({ tabId: message.tabId, message: message.message })
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

const getCoupangProductStocks = async ({ auth, q_change_date, sku_list }) => {
  try {
    const { id_token } = auth

    const stocks = await axios
      .post(
        `${backend.getCoupangProductsStocks}`,
        {
          q_change_date,
          sku_list,
        },
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
        return JSON.parse(res).stocks
      })

    return stocks
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
    const COLUMNS = [
      { header: '스토어ID', key: 'store_id', width: 10 },
      { header: '발주번호', key: 'order_no', width: 10 },
      { header: '물류센터', key: 'logis_center_name', width: 10 },
      { header: '발주상태', key: 'order_status', width: 15 },
      { header: '상품번호', key: 'item_sku_id', width: 10 },
      { header: '상품바코드', key: 'item_sku_barcode', width: 15 },
      { header: '상품명', key: 'item_sku_name', width: 50 },
      { header: '요청수량', key: 'order_qty', width: 10 },
      { header: '납품수량', key: 'confirmed_qty', width: 10 },
      { header: '창고재고', key: 'hq_stock_qty', width: 10 },
      { header: '입고예정', key: 'hq_stock_pending_qty', width: 10 },
      { header: '본사입고예정일', key: 'pending_detail', width: 20 },
      { header: '비교재고', key: 'remaining_hq_stock_qty_after', width: 10 },
      { header: '발주요청일', key: 'order_datetime', width: 10 },
      { header: '쿠팡 입고요청일', key: 'expected_receive_date', width: 15 },
      { header: '옵션ID', key: 'item_id', width: 15 },
    ]
    const headers = COLUMNS.map((column) => column.header)
    const headerWidths = COLUMNS.map((column) => column.width)

    // ExcelJS로 워크북 생성
    const wb = new Excel.Workbook()
    const sheet = wb.addWorksheet('로켓서플라이어 주문')

    // 헤더 추가
    const headerRow = sheet.addRow(headers)
    headerRow.eachCell((cell, colNum) => {
      sheet.getColumn(colNum).width = headerWidths[colNum - 1]
      // cell.font = { bold: true }
      cell.alignment = { vertical: 'middle', horizontal: 'center' }
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'c0c0c0' } }
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      }
    })

    console.log('orders : ', orders)

    orders.forEach((order) => {
      let {
        store_id,
        order_no,
        logis_center_name,
        order_status,
        item_sku_name,
        item_sku_id,
        item_sku_barcode,
        order_qty,
        confirmed_qty,
        hq_stock_qty,
        hq_stock_pending_qty,
        pending_detail,
        remaining_hq_stock_qty_after,
        order_datetime,
        expected_receive_date,
        item_id,
      } = order

      if (pending_detail) {
        const pending_detail_parsed = JSON.parse(pending_detail)

        pending_detail =
          _.chain(pending_detail_parsed)
            .map((item) => `${item?.hq_stock_pending_arrive_date} - ${item?.hq_stock_pending_qty}`)
            .join('\r\n')
            .value() ?? ''
      }

      const rowData = [
        store_id, // 스토어ID
        order_no, // 발주번호
        logis_center_name, // 물류센터
        order_status, // 발주상태
        item_sku_id, // 상품번호
        item_sku_barcode, // 상품바코드
        item_sku_name, // 상품명
        parseInt(order_qty), // 요청수량
        confirmed_qty || '', // 납품수량
        hq_stock_qty || '', // 창고재고
        hq_stock_pending_qty || '', // 입고예정
        pending_detail || '', // 본사입고예정일
        remaining_hq_stock_qty_after || '', // 비교재고
        dayjs(order_datetime).format('YYYY-MM-DD') || '', // 발주요청일
        dayjs(expected_receive_date).format('YYYY-MM-DD') || '', // 쿠팡 입고요청일
        item_id || '', // 옵션ID
      ]
      const row = sheet.addRow(rowData)

      // 모든 셀에 테두리 적용
      row.eachCell((cell, colNum) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        }
      })

      // 납품수량 열(9번째 열)에 노란색 배경 적용
      row.getCell(9).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF00' }, // 노란색
      }

      // 창고재고 열(10번째 열)에 하늘색 배경 적용
      row.getCell(10).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '87CEEB' }, // 하늘색
      }

      // 본사입고예정일 열(12번째 열)에 노란색 배경 적용
      row.getCell(12).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF00' }, // 노란색
      }
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

const downloadCenterExcel = async ({ fcList }) => {
  try {
    const addHeaders = ['후보센터']
    const addHeaderWidths = [50]

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
    })

    // D열에 요청사유 선택 옵션 추가
    const requestReasons = [
      '1.예약 가능한 잔여 슬롯 없음',
      '2.납품 가능한 시간대의 슬롯 없음',
      '3.물류센터 운영 임시 중단',
      '4.밀크런 접수 불가',
      '5.택배사 접수 불가',
      '6.업체 생산 capa 초과',
      '7.MOQ 이하 발주수량',
      '8.물류비 절감을 위해',
      '9.상품 수입 지연',
      '10.재고 부족',
      '11.업체 휴무',
      '12.천재지변',
      '13.정규 발주 입고일로 변경',
      '14.상품 크기, 파렛트 이슈로 입고 불가',
      '15.팔렛단위의 납품 불가(1 SKU = 1Pallet 납품 불가)',
      '16.상품의 유통기한이 100일 이내 상품',
      '17.행사 상품',
      '18.상품 품절 임박/ 품절상태',
    ]
    const listColLetter = 'Z' // 보조 열
    const listColIndex = 26 // Z = 26

    requestReasons.forEach((text, i) => {
      sheet.getCell(i + 1, listColIndex).value = text // Z1, Z2, ...
    })
    sheet.getColumn(listColIndex).hidden = true // 보조 열 숨김

    // 3행부터 데이터 입력
    const sheetName = sheet.name
    fcList.forEach((item, index) => {
      const row = sheet.getRow(3 + index)
      row.getCell(1).value = item.order_no || ''
      row.getCell(2).value = item.center || ''
      row.getCell(3).value = item.edd || ''

      // D열 드롭다운: 같은 시트 범위(Z1:Z18) 참조 (교차 시트 금지!)
      const dCell = row.getCell(4)
      dCell.dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: [
          `='${sheetName}'!$${listColLetter}$1:$${listColLetter}$${requestReasons.length}`,
        ],
        showErrorMessage: true,
        errorStyle: 'warning',
      }

      // 기본 선택값 (목록 항목과 "완전히 동일"해야 함)
      dCell.value = requestReasons[7]

      row.getCell(6).value = item.candidates.map((c) => c.fcName).join(',') || ''
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

const injectAndSendMessage = async ({ tabId, message }) => {
  try {
    console.log('background - injectAndSendMessage 호출:', { tabId, message })

    // 탭 정보 확인
    const tab = await chrome.tabs.get(tabId)
    console.log('background - 탭 정보:', tab.url)

    // content script로 메시지 전송 시도
    try {
      const response = await chrome.tabs.sendMessage(tabId, message)
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

          // 더 긴 대기 시간
          await new Promise((resolve) => setTimeout(resolve, 2000))

          const response = await chrome.tabs.sendMessage(tabId, message)
          console.log('background - 재시도 성공:', response)
          return response
        } catch (injectError) {
          console.error('background - inject 실패:', injectError)
          // 더 자세한 에러 정보
          throw new Error(
            `Content script 로드 실패: ${injectError.message}. 페이지를 새로고침한 후 다시 시도해주세요.`
          )
        }
      }
      throw error
    }
  } catch (error) {
    console.error('background - injectAndSendMessage 오류:', error)
    throw error
  }
}
