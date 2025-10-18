// lodash는 전역으로 사용 (popup에서 이미 로드됨)
// axios는 fetch API로 대체

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log('content script - message received:', message)
  console.log('content script - sender:', sender)

  switch (message.action) {
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
    console.log('body : ', body)
    console.log('q_change_date : ', q_change_date)

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

const sleepRand = (delayMin, delayMax) => {
  const delay = Math.floor(Math.random() * (delayMax - delayMin + 1)) + delayMin
  console.log('대기중 : ', delay)
  return new Promise((resolve) => setTimeout(resolve, delay))
}
