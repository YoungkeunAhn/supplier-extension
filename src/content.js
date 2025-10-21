chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  switch (message.action) {
    case 'getSupplierOrders':
      getProductsPoExpectedReceiveSkuListCoupangRocket(message.q_date)
        .then((centers) => {
          sendResponse(centers)
        })
        .catch((err) => {
          sendResponse({ error: err.message })
        })
      return true

    case 'getCenters':
      getCenters(message.orders, message.q_change_date)
        .then((centers) => {
          sendResponse(centers)
        })
        .catch((err) => {
          sendResponse({ error: err.message })
        })

      return true
    default:
      sendResponse({ error: '알 수 없는 액션입니다.' })
      return true
  }
})

console.log('content script - 로드 완료')

const getProductsPoExpectedReceiveSkuListCoupangRocket = async (q_date) => {
  try {
    let data = []

    let productsPoHeader = ''
    let productsPoText = ''
    let page = 1
    let moreDataAvailable = true
    let after = ''
    const url = 'https://supplier.coupang.com/scm/purchase/order/sku/list-download'
    while (moreDataAvailable) {
      const requestBody = `page=${page}&searchDateType=WAREHOUSING_PLAN_DATE&searchStartDate=${q_date}&searchEndDate=${q_date}&purchaseOrderSeq=&centerCode=&skuSeq=&purchaseOrderStatus=&purchaseOrderType=&selected=[{"id":"purchaseOrderSeq","name":"발주번호"},{"id":"purchaseOrderTypeDesc","name":"발주유형"},{"id":"purchaseOrderStatusDesc","name":"발주현황"},{"id":"skuSeq","name":"SKU ID"},{"id":"skuName","name":"SKU 이름"},{"id":"skuBarcode","name":"SKU Barcode"},{"id":"centerName","name":"물류센터"},{"id":"requiredDeliveryDate","name":"입고예정일"},{"id":"createdAt","name":"발주일"},{"id":"orderedQuantity","name":"발주수량"},{"id":"confirmedQuantity","name":"확정수량"},{"id":"receivedQuantity","name":"입고수량"},{"id":"taxTypePurchaseDesc","name":"매입유형"},{"id":"dutyFreeType","name":"면세여부"},{"id":"producedYear","name":"생산연도"},{"id":"manufacturedDate","name":"제조일자"},{"id":"expirationDate","name":"유통(소비)기한"},{"id":"purchasePrice","name":"매입가"},{"id":"unitPrice","name":"공급가"},{"id":"vatPrice","name":"부가세"},{"id":"orderedAmount","name":"총발주 매입금"},{"id":"receivedAmount","name":"입고금액"}]&size=1000${
        after ? `&after=${after}` : ''
      }`

      const response = await fetch(url, {
        method: 'POST',
        credentials: 'include',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        },
        body: requestBody,
      })

      const body = await response.json()

      if (body?.after) {
        // console.log("page", page)
        page += 1
        after = body?.after
        productsPoHeader = body?.header
        await sleep(3000)
      } else {
        // console.log("end")
        moreDataAvailable = false
      }
      if (body?.data) {
        productsPoText += body.data
      }
    }

    console.log('productsPoHeader : ', productsPoHeader)
    console.log('productsPoText : ', productsPoText)

    const headers = productsPoHeader.split(',')
    const rows = productsPoText.split('\n').filter((row) => row?.trim() !== '')
    const result = rows.map((row) => {
      const values = parseLine(row) //row.split(',');
      const obj = {}
      for (const [index, header] of headers.entries()) {
        obj[header?.trim()] = values[index].trim()
      }
      return obj
    })

    for (const item of result) {
      const new_item = {
        product_type_code: 'RETAIL',
        order_no: item['발주번호'],
        order_type: item['발주유형'],
        order_status: item['발주현황'],
        item_sku_id: item['SKU ID'],
        item_sku_name: item['SKU 이름'],
        item_sku_barcode: item['SKU Barcode'],
        logis_center_name: item['물류센터'],
        expected_receive_date: item['입고예정일'],
        order_datetime: item['발주일'],
        order_qty: item['발주수량'],
        confirmed_qty: item['발주현황'] === '거래처확인요청' ? 0 : item['확정수량'],
        receive_qty: item['입고수량'],
      }
      data.push(new_item)
    }
    console.log('data', data.length)
    console.log('data : ', data)
    return data
  } catch (error) {
    console.log(error)
    throw error
  }
}

const getCenters = async (orders, q_change_date) => {
  let centers = []
  // lodash 대신 vanilla JS로 중복 제거
  const uniqOrders = orders.filter(
    (order, index, self) => index === self.findIndex((o) => o.order_no === order.order_no)
  )
  console.log('uniqOrders : ', uniqOrders)

  for (const order of uniqOrders) {
    const center = await getCenter(order, q_change_date)
    centers.push({ order_no: order.order_no, center })
    await sleepRand(500, 500)
  }

  return centers
}

const getCenter = async (order, q_change_date) => {
  const url = 'https://supplier.coupang.com/plan/v1/ticket/poEditing/getNewCalendarAvailableList'

  // ?poId=112580597&partnerId=A00132210&skuId=&reason=&request_model=&datetime=1760080272971
  const params = new URLSearchParams({
    poId: order.order_no,
    partnerId: order.store_id,
    skuId: '',
    reason: '',
    request_model: '',
    datetime: Date.now(),
  })

  try {
    const response = await fetch(`${url}?${params}`, {
      method: 'GET',
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json',
      },
    })

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`)
    }

    const { body } = await response.json()

    const filteredCenters = body.filter(
      (c) => ['AVAILABLE', 'WARNING'].includes(c.status) && c.edd === q_change_date
    )
    console.log('filteredCenters : ', filteredCenters)
    return filteredCenters
  } catch (error) {
    console.error('getCenter error:', error)
    throw error
  }
}

const sleep = (ms) => {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

const sleepRand = (delayMin, delayMax) => {
  const delay = Math.floor(Math.random() * (delayMax - delayMin + 1)) + delayMin
  console.log('대기중 : ', delay)
  return new Promise((resolve) => setTimeout(resolve, delay))
}

const parseLine = (line) => {
  const result = []
  let current = ''
  let inQuotes = false

  for (let i = 0; i < line.length; i++) {
    const char = line[i]

    if (char === '"') {
      inQuotes = !inQuotes
    } else if (char === ',' && !inQuotes) {
      result.push(current.trim())
      current = ''
    } else {
      current += char
    }
  }

  result.push(current.trim())
  return result
}
