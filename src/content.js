import axios from 'axios'

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  switch (message.action) {
    case 'getCenters':
      getCenters()
        .then((orders) => {
          sendResponse(orders)
        })
        .catch((err) => {
          sendResponse({ error: err.message })
        })

      return true
  }
})

const getCenters = async (orders) => {
  let centers = []
  const uniqOrders = _.uniqBy(orders, 'order_no')

  for (const order of uniqOrders) {
    const center = await getCenter(order)
    await sleepRand()
    centers.push(center)
  }

  return centers
}

const getCenter = async (order) => {
  const url = 'https://supplier.coupang.com/plan/v1/ticket/poEditing/getNewCalendarAvailableList'
  const center = await axios.get(url, {
    params: {
      poId: order.order_no,
      partnerId: order.store_id,
      skuId: '',
      reason: '',
      request_model: '',
      datetime: new Date().now(),
    },
    withCredentials: true,
  })
  return center
}

const sleepRand = () =>
  new Promise((r) =>
    setTimeout(r, Math.floor(Math.random() * (delayMax - delayMin + 1)) + delayMin)
  )
