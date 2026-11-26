import http from 'k6/http';
import { check, sleep, group } from 'k6';
import { Trend, Rate, Counter } from 'k6/metrics';

// ============================================================================
// SharePoint Clone — Performance & Stress Test Suite (k6)
// Endpoints: Health, Default Context, Library Items, Search
// ============================================================================

// Custom metrics for granular reporting
const healthLatency = new Trend('sp_health_duration', true);
const contextLatency = new Trend('sp_context_duration', true);
const libraryItemsLatency = new Trend('sp_library_items_duration', true);
const searchLatency = new Trend('sp_search_duration', true);
const errorRate = new Rate('sp_error_rate');
const requestCounter = new Counter('sp_total_requests');

// Base URL configuration (fallback to local Docker container)
const BASE_URL = __ENV.API_URL || 'http://localhost:5000/api/v1';
const FALLBACK_LIBRARY_ID = '22222222-2222-2222-2222-222222222222';

export const options = {
  stages: [
    { duration: '10s', target: 10 }, // Ramp-up to 10 VUs
    { duration: '30s', target: 20 }, // Steady load at 20 VUs
    { duration: '10s', target: 0 },  // Ramp-down to 0 VUs
  ],
  thresholds: {
    // Mandated requirement: 95th percentile under 500ms
    http_req_duration: ['p(95)<500'],
    http_req_failed: ['rate<0.01'],
    sp_error_rate: ['rate<0.02'],
    sp_health_duration: ['p(95)<200'],
    sp_library_items_duration: ['p(95)<500'],
    sp_search_duration: ['p(95)<500'],
  },
};

// Global setup: fetch default context once to obtain dynamic library ID
export function setup() {
  const res = http.get(`${BASE_URL}/default-context`, {
    headers: { 'Accept': 'application/json' },
  });

  let libraryId = FALLBACK_LIBRARY_ID;
  if (res.status === 200) {
    try {
      const data = JSON.parse(res.body);
      if (data && data.library && data.library.id) {
        libraryId = data.library.id;
      }
    } catch (e) {
      // JSON parse error, use fallback
    }
  }

  return { libraryId };
}

export default function (data) {
  const libraryId = (data && data.libraryId) ? data.libraryId : FALLBACK_LIBRARY_ID;
  const commonHeaders = {
    'Accept': 'application/json',
    'User-Agent': 'k6-sharepoint-load-tester/1.0',
  };

  // 1. Scenario: Health Check
  group('01_Health_Check', function () {
    const res = http.get(`${BASE_URL}/health`, { headers: commonHeaders });
    requestCounter.add(1);
    healthLatency.add(res.timings.duration);

    const isOk = check(res, {
      'health status is 200': (r) => r.status === 200,
      'health body is valid': (r) => {
        try {
          const body = JSON.parse(r.body);
          return body.status === 'healthy' || body.status === 'ok';
        } catch (_) {
          return false;
        }
      },
    });

    errorRate.add(!isOk);
  });

  sleep(0.3);

  // 2. Scenario: Default Context (Site & Library metadata)
  group('02_Default_Context', function () {
    const res = http.get(`${BASE_URL}/default-context`, { headers: commonHeaders });
    requestCounter.add(1);
    contextLatency.add(res.timings.duration);

    const isOk = check(res, {
      'context status is 200': (r) => r.status === 200,
      'context contains site and library': (r) => {
        try {
          const body = JSON.parse(r.body);
          return !!(body.site && body.library);
        } catch (_) {
          return false;
        }
      },
    });

    errorRate.add(!isOk);
  });

  sleep(0.3);

  // 3. Scenario: Library Items Navigation
  group('03_Library_Items', function () {
    const res = http.get(`${BASE_URL}/libraries/${libraryId}/items`, { headers: commonHeaders });
    requestCounter.add(1);
    libraryItemsLatency.add(res.timings.duration);

    const isOk = check(res, {
      'library items status is 200': (r) => r.status === 200,
      'library items payload is array or object': (r) => {
        try {
          const body = JSON.parse(r.body);
          const items = Array.isArray(body) ? body : body.items;
          return Array.isArray(items);
        } catch (_) {
          return false;
        }
      },
    });

    errorRate.add(!isOk);
  });

  sleep(0.3);

  // 4. Scenario: Instant Search
  group('04_Instant_Search', function () {
    const res = http.get(`${BASE_URL}/search?q=a`, { headers: commonHeaders });
    requestCounter.add(1);
    searchLatency.add(res.timings.duration);

    const isOk = check(res, {
      'search status is 200': (r) => r.status === 200,
      'search results returned': (r) => {
        try {
          const body = JSON.parse(r.body);
          return Array.isArray(body) || (body && Array.isArray(body.results));
        } catch (_) {
          return false;
        }
      },
    });

    errorRate.add(!isOk);
  });

  // Jittered think time to simulate realistic user behavior
  sleep(0.5 + Math.random() * 0.5);
}

// Summary formatting
export function handleSummary(data) {
  const p95 = data.metrics.http_req_duration && data.metrics.http_req_duration.values
    ? data.metrics.http_req_duration.values['p(95)'].toFixed(2)
    : 'N/A';
  const totalReqs = data.metrics.sp_total_requests ? data.metrics.sp_total_requests.values.count : 'N/A';
  const failureRate = data.metrics.http_req_failed && data.metrics.http_req_failed.values
    ? (data.metrics.http_req_failed.values.rate * 100).toFixed(2)
    : '0';

  console.log(`\n======================================================`);
  console.log(`🚀 SharePoint Clone — Performance Test Summary (k6)`);
  console.log(`------------------------------------------------------`);
  console.log(`Total Requests:      ${totalReqs}`);
  console.log(`p(95) Duration:      ${p95} ms (Threshold: < 500 ms)`);
  console.log(`Failure Rate:        ${failureRate}%`);
  console.log(`======================================================\n`);

  return {
    stdout: JSON.stringify(data, null, 2),
  };
}
