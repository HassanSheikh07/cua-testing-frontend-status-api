const { app } = require('@azure/functions')

const statusStore = new Map()

function getCorsHeaders () {
  return {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  }
}

function jsonResponse (statusCode, body) {
  return {
    status: statusCode,
    headers: getCorsHeaders(),
    body: JSON.stringify(body)
  }
}

app.http('status', {
  methods: ['GET', 'POST', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'status',
  handler: async function (request, context) {
    try {
      if (request.method === 'OPTIONS') {
        return {
          status: 204,
          headers: getCorsHeaders()
        }
      }

      if (request.method === 'POST') {
        let body = {}

        try {
          body = await request.json()
        } catch (error) {
          return jsonResponse(400, {
            success: false,
            message: 'Invalid JSON body.',
            error: String(error)
          })
        }

        const recordId =
          body.recordId ||
          body.cr720_uahtestscriptid ||
          body.dataverseRecordId ||
          ''

        if (!recordId) {
          return jsonResponse(400, {
            success: false,
            message: 'recordId is required.'
          })
        }

        const record = {
          recordId: recordId,
          scriptName: body.scriptName || '',
          status: body.status || body.cr720_status || 'Unknown',
          rawStatus: body.rawStatus || '',
          stage: body.stage || 'Dataverse',
          message: body.message || '',
          source: body.source || 'Power Automate',
          updatedAt: new Date().toISOString(),
          raw: body
        }

        statusStore.set(recordId, record)

        context.log('Status updated for record:', recordId, record)

        return jsonResponse(200, {
          success: true,
          message: 'Status updated successfully.',
          data: record
        })
      }

      if (request.method === 'GET') {
        const url = new URL(request.url)
        const recordId = url.searchParams.get('recordId')

        if (!recordId) {
          return jsonResponse(400, {
            success: false,
            message: 'recordId query parameter is required.'
          })
        }

        const record = statusStore.get(recordId)

        if (!record) {
          return jsonResponse(404, {
            success: false,
            message: 'No status found for recordId: ' + recordId
          })
        }

        return jsonResponse(200, {
          success: true,
          data: record
        })
      }

      return jsonResponse(405, {
        success: false,
        message: 'Method not allowed.'
      })
    } catch (error) {
      context.error('Status API error:', error)

      return jsonResponse(500, {
        success: false,
        message: 'Internal server error.',
        error: String(error)
      })
    }
  }
})