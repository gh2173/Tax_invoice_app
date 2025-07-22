/*
 * EZVoucher2.js - 매입송장 처리 RPA 자동화
 * 
 * 동작 순서:
 * 1. ERP 접속 및 로그인 
 * 완료
 *    - D365 페이지 접속 (https://d365.nepes.co.kr/namespaces/AXSF/?cmp=K02&mi=DefaultDashboard)
 *    - ADFS 로그인 처리 (#userNameInput, #passwordInput, #submitButton)
 *    - 페이지 로딩 완료 대기
 * 
 * 2. 검색 기능을 통한 구매 입고내역 조회 페이지 이동
 *    - 검색 버튼 클릭 (Find-symbol 버튼)
 *    - "구매 입고내역 조회(N)" 검색어 입력
 *    - NavigationSearchBox에서 해당 메뉴 클릭
 * 
 * 3. (추후 구현 예정) 매입송장 처리 로직
 *    - 파일 업로드
 *    - 데이터 처리
 *    - 결과 확인
 */

const puppeteer = require('puppeteer');
const puppeteerExtra = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const winston = require('winston');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx'); // 엑셀 파일 읽기용 라이브러리

const { ipcMain, dialog } = require('electron');

// 기본 대기 함수
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// 성능 최적화를 위한 스마트 대기 시스템
const smartWait = {
  // 요소가 나타날 때까지 최대 timeout까지 대기
  forElement: async (page, selector, timeout = 5000) => {
    try {
      await page.waitForSelector(selector, { visible: true, timeout });
      return true;
    } catch (error) {
      logger.warn(`요소 대기 시간 초과: ${selector} (${timeout}ms)`);
      return false;
    }
  },

  // 요소가 클릭 가능해질 때까지 대기
  forClickable: async (page, selector, timeout = 5000) => {
    try {
      await page.waitForSelector(selector, { visible: true, timeout });
      await page.waitForFunction(
        (sel) => {
          const el = document.querySelector(sel);
          return el && !el.disabled && el.offsetParent !== null;
        },
        { timeout: 3000 },
        selector
      );
      return true;
    } catch (error) {
      logger.warn(`클릭 가능한 요소 대기 시간 초과: ${selector}`);
      return false;
    }
  },

  // 페이지가 준비될 때까지 대기
  forPageReady: async (page, timeout = 8000) => {
    try {
      await page.waitForFunction(
        () => document.readyState === 'complete',
        { timeout }
      );
      await delay(500); // 추가 안정화 대기
      return true;
    } catch (error) {
      logger.warn(`페이지 준비 대기 시간 초과: ${timeout}ms`);
      return false;
    }
  },

  // 여러 선택자 중 하나가 나타날 때까지 대기
  forAnyElement: async (page, selectors, timeout = 5000) => {
    try {
      await Promise.race(
        selectors.map(selector => 
          page.waitForSelector(selector, { visible: true, timeout })
        )
      );
      return true;
    } catch (error) {
      logger.warn(`복수 요소 대기 시간 초과: ${selectors.join(', ')}`);
      return false;
    }  }
};

/**
 * 데이터 테이블이 로드될 때까지 대기하는 함수
 * @param {Object} page - Puppeteer page 객체
 * @param {number} timeout - 최대 대기 시간 (기본값: 30초)
 * @returns {boolean} - 데이터 테이블이 로드되었는지 여부
 */
async function waitForDataTable(page, timeout = 30000) {
  const startTime = Date.now();
  logger.info(`데이터 테이블 로딩 대기 시작 (최대 ${timeout/1000}초)`);
  
  let loadingCompleted = false;
  
  while (Date.now() - startTime < timeout) {
    try {
      // 1. 로딩 스피너 확인 (있으면 계속 대기)
      const isLoading = await page.evaluate(() => {
        const loadingSelectors = [
          '.loading', '.spinner', '.ms-Spinner', '[aria-label*="로딩"]',
          '[aria-label*="Loading"]', '.dyn-loading', '.loadingSpinner'
        ];
        
        return loadingSelectors.some(selector => {
          const element = document.querySelector(selector);
          if (element) {
            const style = window.getComputedStyle(element);
            return style.display !== 'none' && 
                   style.visibility !== 'hidden' && 
                   element.offsetParent !== null;
          }
          return false;
        });
      });
      
      if (isLoading) {
        logger.info('로딩 중입니다. 계속 대기...');
        loadingCompleted = false; // 로딩이 다시 시작되면 플래그 리셋
        await delay(2000);
        continue;
      }
      
      // 2. 로딩 스피너가 사라진 후 처음이면 10초 대기
      if (!loadingCompleted) {
        logger.info('✅ 로딩 스피너가 사라졌습니다. 안정화를 위해 10초 대기 중...');
        await delay(10000);
        loadingCompleted = true;
        logger.info('안정화 대기 완료. 데이터 그리드 확인 중...');
      }
      
      // 3. 데이터 그리드 확인
      const hasDataGrid = await page.evaluate(() => {
        const gridSelectors = [
          '[data-dyn-controlname*="Grid"]', '.dyn-grid', 'div[role="grid"]',
          'table[role="grid"]', '[class*="grid"]', 'table'
        ];
        
        for (const selector of gridSelectors) {
          const element = document.querySelector(selector);
          if (element) {
            const rows = element.querySelectorAll('tr, [role="row"], [data-dyn-row]');
            if (rows.length > 0) { // 최소 1개 행이 있으면 OK
              return true;
            }
          }
        }
        return false;
      });
      
      if (hasDataGrid) {
        logger.info('✅ 데이터 그리드가 감지되었습니다. 테이블 로딩 완료!');
        return true;
      }
      
      logger.info('데이터 그리드를 찾는 중...');
      await delay(2000);
      
    } catch (error) {
      logger.warn(`데이터 테이블 대기 중 오류: ${error.message}`);
      await delay(2000);
    }
  }
  
  logger.warn(`⚠️ 데이터 테이블 로딩 대기 시간 초과 (${timeout/1000}초)`);
  return false;
}

// 로거 설정
const logger = winston.createLogger({
  level: 'info',
  format: winston.format.combine(
    winston.format.timestamp({
      format: 'YYYY-MM-DD HH:mm:ss'
    }),
    winston.format.errors({ stack: true }),
    winston.format.json()
  ),
  defaultMeta: { service: 'EZVoucher2-RPA' },
  transports: [
    new winston.transports.File({ 
      filename: path.join(__dirname, 'rpa.log'),
      maxsize: 5242880, // 5MB
      maxFiles: 5
    }),
    new winston.transports.Console({
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.simple()
      )
    })
  ]
});

// Puppeteer Extra 설정
puppeteerExtra.use(StealthPlugin());

// 글로벌 변수
let globalCredentials = {
  username: '',
  password: ''
};

// 로그인 처리 함수 (EZVoucher.js와 동일한 ADFS 전용 로직)
async function handleLogin(page, credentials) {
  try {
    // 1. 사용자 이름(이메일) 입력
    logger.info('사용자 이름 입력 중...');
    await page.waitForSelector('#userNameInput', { visible: true, timeout: 10000 });
    await page.type('#userNameInput', credentials.username);
    logger.info('사용자 이름 입력 완료');
    
    // 2. 비밀번호 입력
    logger.info('비밀번호 입력 중...');
    await page.waitForSelector('#passwordInput', { visible: true, timeout: 10000 });
    await page.type('#passwordInput', credentials.password);
    logger.info('비밀번호 입력 완료');
    
    // 3. 로그인 버튼 클릭
    logger.info('로그인 버튼 클릭 중...');
    await page.waitForSelector('#submitButton', { visible: true, timeout: 10000 });
    await page.click('#submitButton');
    logger.info('로그인 버튼 클릭 완료');
    
    // 로그인 후 페이지 로드 대기
    logger.info('로그인 후 페이지 로드 대기 중...');
    await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 });
    
    // 로그인 성공 확인
    logger.info('로그인 완료');
    
  } catch (error) {
    // 오류 시 스크린샷
    logger.error(`로그인 오류: ${error.message}`);
    throw error;
  }
}

// 글로벌 로그인 정보 설정
function setCredentials(username, password) {
  globalCredentials.username = username;
  globalCredentials.password = password;
  logger.info('매입송장 처리용 로그인 정보가 설정되었습니다');
}

// 글로벌 로그인 정보 반환
function getCredentials() {
  return globalCredentials;
}

// 1. ERP 접속 및 로그인 완료
async function connectToD365(credentials) {
  logger.info('=== 매입송장 처리 - D365 접속 시작 ===');
  
  const browser = await puppeteerExtra.launch({
    headless: false,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-web-security',
      '--disable-features=VizDisplayCompositor',
      '--start-maximized',
      '--ignore-certificate-errors',
      '--ignore-ssl-errors',
      '--ignore-certificate-errors-spki-list'
    ],
    defaultViewport: null
  });

  const page = await browser.newPage();
  
  try {
    // User-Agent 설정
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
    
    // SSL 인증서 오류 처리
    await page.setBypassCSP(true);
    
    // 페이지 요청 인터셉트 설정 (SSL 오류 처리용)
    await page.setRequestInterception(true);
    page.on('request', request => {
      request.continue();
    });
    
    // 대화상자 처리 (인증서 경고 등)
    page.on('dialog', async dialog => {
      logger.info(`대화상자 감지: ${dialog.message()}`);
      await dialog.accept();
    });
    
    // D365 페이지 접속 (재시도 로직 추가)
    logger.info('D365 페이지로 이동 중...');
    let pageLoadSuccess = false;
    let retryCount = 0;
    const maxRetries = 3;
    
    while (!pageLoadSuccess && retryCount < maxRetries) {
      try {
        retryCount++;
        logger.info(`D365 페이지 접속 시도 ${retryCount}/${maxRetries}`);
        
        await page.goto('https://d365.nepes.co.kr/namespaces/AXSF/?cmp=K02&mi=DefaultDashboard', {
          waitUntil: 'networkidle2',
          timeout: 60000 // 60초 타임아웃
        });
        
        pageLoadSuccess = true;
        logger.info('D365 페이지 로드 완료');
      } catch (networkError) {
        logger.error(`D365 페이지 접속 시도 ${retryCount} 실패: ${networkError.message}`);
        
        if (retryCount >= maxRetries) {
          const errorMsg = `네트워크 연결 실패: D365 사이트(https://d365.nepes.co.kr)에 접속할 수 없습니다. 인터넷 연결을 확인하거나 VPN이 필요할 수 있습니다.`;
          logger.error(errorMsg);
          throw new Error(errorMsg);
        }
        
        // 재시도 전 2초 대기
        logger.info('2초 후 재시도합니다...');
        await delay(2000);
      }
    }    // 로그인 처리 (필요한 경우) - EZVoucher.js와 동일한 조건
    if (await page.$('input[type="email"]') !== null || await page.$('#userNameInput') !== null) {
      logger.info('로그인 화면 감지됨, 로그인 시도 중...');
      await handleLogin(page, credentials);
    }
    
    // 로그인 후 페이지가 완전히 로드될 때까지 스마트 대기
    logger.info('로그인 후 페이지 로딩 확인 중...');
    const pageReady = await smartWait.forPageReady(page, 8000);
    if (!pageReady) {
      logger.warn('페이지 로딩 확인 실패, 기본 2초 대기로 진행');
      await delay(2000);
    }    logger.info('페이지 로딩 확인 완료');
    
    logger.info('=== 1. ERP 접속 및 로그인 완료 ===');
    
    // 2번 동작 실행: 구매 입고내역 조회 페이지 이동
    await navigateToReceivingInquiry(page);
    
    // 3번 동작 실행: 엑셀 파일 열기 및 매크로 실행 (page 매개변수 전달)
    logger.info('🚀 === 3번 동작: 엑셀 파일 열기 및 매크로 실행 시작 ===');
    const excelResult = await executeExcelProcessing(page);
    if (!excelResult.success) {
      logger.warn(`⚠️ 엑셀 처리 실패: ${excelResult.error}`);
    } else {
      logger.info('✅ 3번 동작: 엑셀 파일 열기 및 매크로 실행 완료');
      logger.info('✅ 4번 동작: 대기중인 공급사송장 메뉴 이동도 완료됨');
    }
    
    
    // 전체 프로세스 완료 대기
    await delay(5000);
    
    // 완료 팝업창 표시
    try {
      await page.evaluate(() => {
        alert('🎉 매입송장 처리 RPA 자동화가 완료되었습니다!\n\n✅ 1. ERP 접속 및 로그인 완료\n✅ 2. 구매 입고내역 조회 및 다운로드 완료\n✅ 3. 엑셀 파일 열기 및 매크로 실행 완료\n✅ 4. 대기중인 공급사송장 메뉴 이동 및 엑셀 데이터 처리 완료\n\n브라우저를 직접 닫아주세요.');
      });
      logger.info('✅ 완료 팝업창 표시됨');
    } catch (alertError) {
      logger.warn(`완료 팝업창 표시 실패: ${alertError.message}`);
    }
    
    logger.info('🎉 === 전체 RPA 프로세스 완료 - 브라우저는 열린 상태로 유지됩니다 ===');
      // 성공 시 serializable한 객체만 반환
    return { 
      success: true, 
      message: '1. ERP 접속 및 로그인 완료\n2. 구매 입고내역 조회 및 다운로드 완료\n3. 엑셀 파일 열기 및 매크로 실행 완료\n4. 대기중인 공급사송장 메뉴 이동 및 엑셀 데이터 처리 완료',
      completedAt: new Date().toISOString(),
      browserKeptOpen: true
    };
    
  } catch (error) {
    logger.error(`D365 접속 중 오류 발생: ${error.message}`);
    
    // 에러 팝업창 표시
    try {
      await page.evaluate((errorMsg) => {
        alert(`❌ 매입송장 처리 RPA 자동화 중 오류가 발생했습니다!\n\n오류 내용: ${errorMsg}\n\n브라우저를 직접 닫아주세요.`);
      }, error.message);
      logger.info('❌ 에러 팝업창 표시됨');
    } catch (alertError) {
      logger.warn(`에러 팝업창 표시 실패: ${alertError.message}`);
    }
    
    // 에러 시에도 serializable한 객체만 반환
    return { 
      success: false, 
      error: error.message,
      failedAt: new Date().toISOString(),
      browserKeptOpen: true
    };
  } finally {
    // finally 블록에서 브라우저를 닫지 않음 - 사용자가 결과를 확인할 수 있도록 유지
    logger.info('브라우저는 열린 상태로 유지됩니다. 사용자가 직접 닫아주세요.');
  }
}

// 2. 검색 기능을 통한 구매 입고내역 조회 페이지 이동
async function navigateToReceivingInquiry(page) {
  logger.info('=== 2. 구매 입고내역 조회 페이지 이동 시작 ===');
  
  try {
    // 2-1. 검색 버튼 클릭 (Find-symbol 버튼)
    logger.info('검색 버튼 찾는 중...');
    
    const searchButtonSelectors = [
      '.button-commandRing.Find-symbol',
      'span.Find-symbol',
      '[data-dyn-image-type="Symbol"].Find-symbol',
      '.button-container .Find-symbol'
    ];
    
    let searchButtonClicked = false;
    
    for (const selector of searchButtonSelectors) {
      try {
        logger.info(`검색 버튼 선택자 시도: ${selector}`);
        
        const searchButton = await page.$(selector);
        if (searchButton) {
          const isVisible = await page.evaluate(el => {
            const style = window.getComputedStyle(el);
            return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
          }, searchButton);
          
          if (isVisible) {
            await searchButton.click();
            logger.info(`검색 버튼 클릭 성공: ${selector}`);
            searchButtonClicked = true;
            break;
          } else {
            logger.warn(`검색 버튼이 보이지 않음: ${selector}`);
          }
        }
      } catch (error) {
        logger.warn(`검색 버튼 클릭 실패: ${selector} - ${error.message}`);
      }
    }
    
    if (!searchButtonClicked) {
      // JavaScript로 직접 검색 버튼 클릭 시도
      try {
        logger.info('JavaScript로 검색 버튼 직접 클릭 시도...');
        await page.evaluate(() => {
          const searchButtons = document.querySelectorAll('.Find-symbol, [data-dyn-image-type="Symbol"]');
          for (const btn of searchButtons) {
            if (btn.classList.contains('Find-symbol') || btn.getAttribute('data-dyn-image-type') === 'Symbol') {
              btn.click();
              return true;
            }
          }
          return false;
        });
        searchButtonClicked = true;
        logger.info('JavaScript로 검색 버튼 클릭 성공');
      } catch (jsError) {
        logger.error('JavaScript 검색 버튼 클릭 실패:', jsError.message);
      }
    }
    
    if (!searchButtonClicked) {
      throw new Error('검색 버튼을 찾을 수 없습니다.');
    }
    
    // 검색창이 나타날 때까지 대기
    await delay(2000);
    
    // 2-2. "구매 입고내역 조회(N)" 검색어 입력
    logger.info('검색어 입력 중...');
    
    const searchInputSelectors = [
      'input[type="text"]',
      '.navigationSearchBox input',
      '#NavigationSearchBox',
      'input[placeholder*="검색"]',
      'input[aria-label*="검색"]'
    ];
    
    let searchInputFound = false;
    const searchTerm = '구매 입고내역 조회(N)';
    
    for (const selector of searchInputSelectors) {
      try {
        logger.info(`검색 입력창 선택자 시도: ${selector}`);
        
        await page.waitForSelector(selector, { visible: true, timeout: 5000 });
        
        // 기존 텍스트 클리어
        await page.click(selector, { clickCount: 3 }); // 모든 텍스트 선택
        await page.keyboard.press('Backspace'); // 선택된 텍스트 삭제
        
        // 검색어 입력
        await page.type(selector, searchTerm, { delay: 100 });
        logger.info(`검색어 입력 완료: ${searchTerm}`);
        
        searchInputFound = true;
        break;
        
      } catch (error) {
        logger.warn(`검색 입력창 처리 실패: ${selector} - ${error.message}`);
      }
    }
    
    if (!searchInputFound) {
      throw new Error('검색 입력창을 찾을 수 없습니다.');
    }
    
    // 검색 결과가 나타날 때까지 대기
    await delay(3000);
    
    // 2-3. NavigationSearchBox에서 해당 메뉴 클릭
    logger.info('검색 결과에서 구매 입고내역 조회 메뉴 찾는 중...');
    
    const searchResultSelectors = [
      '.navigationSearchBox',
      '.search-results',
      '.navigation-search-results',
      '[data-dyn-bind*="NavigationSearch"]'
    ];
    
    let menuClicked = false;
    
    for (const containerSelector of searchResultSelectors) {
      try {
        const container = await page.$(containerSelector);
        if (container) {
          // 컨테이너 내에서 "구매 입고내역 조회" 텍스트가 포함된 요소 찾기
          const menuItems = await page.$$eval(`${containerSelector} *`, (elements) => {
            return elements
              .filter(el => {
                const text = el.textContent || el.innerText || '';
                return text.includes('구매 입고내역 조회') || text.includes('구매') && text.includes('입고');
              })
              .map(el => ({
                text: el.textContent || el.innerText,
                clickable: el.tagName === 'A' || el.tagName === 'BUTTON' || el.onclick || el.getAttribute('role') === 'button'
              }));
          });
          
          logger.info(`검색 결과 메뉴 항목들:`, menuItems);
          
          if (menuItems.length > 0) {
            // 첫 번째 매칭되는 항목 클릭
            await page.evaluate((containerSel) => {
              const container = document.querySelector(containerSel);
              if (container) {
                const elements = container.querySelectorAll('*');
                for (const el of elements) {
                  const text = el.textContent || el.innerText || '';
                  if (text.includes('구매 입고내역 조회') || (text.includes('구매') && text.includes('입고'))) {
                    el.click();
                    return true;
                  }
                }
              }
              return false;
            }, containerSelector);
            
            logger.info('구매 입고내역 조회 메뉴 클릭 완료');
            menuClicked = true;
            break;
          }
        }
      } catch (error) {
        logger.warn(`검색 결과 처리 실패: ${containerSelector} - ${error.message}`);
      }
    }
    
    if (!menuClicked) {
      // Enter 키로 첫 번째 결과 선택 시도
      logger.info('Enter 키로 검색 결과 선택 시도...');
      await page.keyboard.press('Enter');
      menuClicked = true;
    }
    
    // 페이지 이동 대기
    logger.info('구매 입고내역 조회 페이지 로딩 대기 중...');
    await delay(5000);
    
    // 페이지 로딩 완료 확인
    const pageReady = await smartWait.forPageReady(page, 10000);
    if (!pageReady) {
      logger.warn('페이지 로딩 확인 실패, 기본 3초 대기로 진행');
      await delay(3000);
    }
    
    logger.info('=== 2. 구매 입고내역 조회 페이지 이동 완료 ===');


    
    // 3. FromDate 입력 (현재 월의 첫날)
    logger.info('=== 3. FromDate 설정 시작 ===');
    
    // 현재 날짜에서 월의 첫날 계산
    const now = new Date();
    // 현재날짜 기준 현재월 가져오기
    const fromDate = `${now.getMonth() + 1}/1/${now.getFullYear()}`; // M/d/YYYY 형태

    logger.info(`설정할 FromDate: ${fromDate}`);
    
    
    // Test 임시 수정
    /*// 현재 날짜에서 지난 월의 첫날 계산
    const now = new Date();
    // 지난달 계산 (현재 월에서 1을 뺌)
    const lastMonth = now.getMonth();
    const year = lastMonth === 0 ? now.getFullYear() -1 : now.getFullYear();
    const month = lastMonth === 0 ? 12 : lastMonth;

    // 지난달 첫날 설정
    const fromDate = `${month}/1/${year}`; // M/d/YYYY 형태

    logger.info(`설정할 FromDate: ${fromDate}`);
    */
    //-------------------------------------------------------------------------------

    // FromDate 입력창 선택자들
    const fromDateSelectors = [
      'input[name="FromDate"]',
      'input[id*="FromDate_input"]',
      'input[aria-labelledby*="FromDate_label"]',
      'input[placeholder=""][name="FromDate"]'
    ];
    
    let fromDateSet = false;
    
    for (const selector of fromDateSelectors) {
      try {
        logger.info(`FromDate 입력창 선택자 시도: ${selector}`);
        
        await page.waitForSelector(selector, { visible: true, timeout: 5000 });
        
        // 입력창 클릭
        await page.click(selector);
        await delay(500);
        
        // 기존 텍스트 클리어 (모든 텍스트 선택 후 삭제)
        await page.click(selector, { clickCount: 3 });
        await page.keyboard.press('Backspace');
        await delay(300);
        
        // 날짜 입력
        await page.type(selector, fromDate, { delay: 100 });
        await page.keyboard.press('Tab'); // 포커스 이동으로 입력 확정
        
        logger.info(`FromDate 설정 완료: ${fromDate}`);
        fromDateSet = true;
        break;
        
      } catch (error) {
        logger.warn(`FromDate 설정 실패: ${selector} - ${error.message}`);
      }
    }
    
    if (!fromDateSet) {
      throw new Error('FromDate 입력창을 찾을 수 없습니다.');
    }
    
    await delay(1000); // 입력 안정화 대기
    
    // 4. ToDate 입력 (현재 월의 마지막 날)
    logger.info('=== 4. ToDate 설정 시작 ===');
    
    
    // 현재 날짜에서 월의 마지막 날 계산
    const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
    const toDate = `${now.getMonth() + 1}/${lastDay}/${now.getFullYear()}`; // M/d/YYYY 형태
    logger.info(`설정할 ToDate: ${toDate}`);
    
    // Test 임시 수정
    // 지난달의 마지막 날 계산
    /*const lastDay = new Date(year, month, 0).getDate();
    const toDate = `${month}/${lastDay}/${year}`; // M/d/YYYY 형태
    logger.info(`설정할 ToDate: ${toDate}`);
    */
    // ToDate 입력창 선택자들
    const toDateSelectors = [
      'input[name="ToDate"]',
      'input[id*="ToDate_input"]',
      'input[aria-labelledby*="ToDate_label"]',
      'input[placeholder=""][name="ToDate"]'
    ];
    
    let toDateSet = false;
    
    for (const selector of toDateSelectors) {
      try {
        logger.info(`ToDate 입력창 선택자 시도: ${selector}`);
        
        await page.waitForSelector(selector, { visible: true, timeout: 5000 });
        
        // 입력창 클릭
        await page.click(selector);
        await delay(500);
        
        // 기존 텍스트 클리어 (모든 텍스트 선택 후 삭제)
        await page.click(selector, { clickCount: 3 });
        await page.keyboard.press('Backspace');
        await delay(300);
        
        // 날짜 입력
        await page.type(selector, toDate, { delay: 100 });
        await page.keyboard.press('Tab'); // 포커스 이동으로 입력 확정
        
        logger.info(`ToDate 설정 완료: ${toDate}`);
        toDateSet = true;
        break;
        
      } catch (error) {
        logger.warn(`ToDate 설정 실패: ${selector} - ${error.message}`);
      }
    }
    
    if (!toDateSet) {
      throw new Error('ToDate 입력창을 찾을 수 없습니다.');
    }
    
    await delay(1000); // 입력 안정화 대기
    
    // 5. Inquiry 버튼 클릭
    logger.info('=== 5. Inquiry 버튼 클릭 시작 ===');
    
    // Inquiry 버튼 선택자들
    const inquiryButtonSelectors = [
      '.button-container:has(.button-label:contains("Inquiry"))',
      'span.button-label:contains("Inquiry")',
      'div.button-container span[id*="Inquiry_label"]',
      '[id*="Inquiry_label"]',
      'span[for*="Inquiry"]'
    ];
    
    let inquiryButtonClicked = false;
    
    for (const selector of inquiryButtonSelectors) {
      try {
        logger.info(`Inquiry 버튼 선택자 시도: ${selector}`);
        
        // CSS 선택자에 :contains()가 있는 경우 JavaScript로 처리
        if (selector.includes(':contains(')) {
          const clicked = await page.evaluate(() => {
            const buttons = document.querySelectorAll('.button-container');
            for (const container of buttons) {
              const label = container.querySelector('.button-label, span[id*="label"]');
              if (label && (label.textContent || label.innerText || '').includes('Inquiry')) {
                container.click();
                return true;
              }
            }
            return false;
          });
          
          if (clicked) {
            logger.info('JavaScript로 Inquiry 버튼 클릭 성공');
            inquiryButtonClicked = true;
            break;
          }
        } else {
          // 일반 선택자 처리
          const inquiryButton = await page.$(selector);
          if (inquiryButton) {
            const isVisible = await page.evaluate(el => {
              const style = window.getComputedStyle(el);
              return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
            }, inquiryButton);
            
            if (isVisible) {
              await inquiryButton.click();
              logger.info(`Inquiry 버튼 클릭 성공: ${selector}`);
              inquiryButtonClicked = true;
              break;
            } else {
              logger.warn(`Inquiry 버튼이 보이지 않음: ${selector}`);
            }
          }
        }
      } catch (error) {
        logger.warn(`Inquiry 버튼 클릭 실패: ${selector} - ${error.message}`);
      }
    }
    
    // 추가 시도: ID와 텍스트를 조합한 방법
    if (!inquiryButtonClicked) {
      try {
        logger.info('ID와 텍스트 조합으로 Inquiry 버튼 찾는 중...');
        
        const clicked = await page.evaluate(() => {
          // id에 "Inquiry"가 포함된 요소들 찾기
          const elements = document.querySelectorAll('[id*="Inquiry"]');
          for (const el of elements) {
            // 클릭 가능한 요소이거나 부모가 클릭 가능한 요소인지 확인
            const clickableEl = el.closest('.button-container, button, [role="button"]') || el;
            if (clickableEl) {
              clickableEl.click();
              return true;
            }
          }
          return false;
        });
        
        if (clicked) {
          logger.info('ID 기반으로 Inquiry 버튼 클릭 성공');
          inquiryButtonClicked = true;
        }
      } catch (error) {
        logger.warn(`ID 기반 Inquiry 버튼 클릭 실패: ${error.message}`);
      }
    }
      if (!inquiryButtonClicked) {
      throw new Error('Inquiry 버튼을 찾을 수 없습니다.');
    }
    
    // 조회 실행 후 데이터 테이블이 나타날 때까지 대기
    logger.info('조회 실행 중, 데이터 테이블 로딩 대기...');
    
    // 기본 대기 시간 (최소 10초 - 조회 실행 후 초기 로딩 대기)
    await delay(10000);
    
    // 데이터 테이블 로딩 확인 (30초 타임아웃으로 단축)
    const dataTableLoaded = await waitForDataTable(page, 30000);
    
    if (!dataTableLoaded) {
      logger.warn('데이터 테이블 로딩 확인 실패, 하지만 계속 진행합니다...');
      // 추가 대기 후 계속 진행
      await delay(5000);
    }
      logger.info('=== 구매 입고내역 조회 설정 및 조회 실행 완료 ===');
    
    // 6. 데이터 내보내기 실행
    logger.info('🚀 === 6. 데이터 내보내기 시작 ===');
    
    // 내보내기 전 추가 안정화 대기
    await delay(3000);
    
    // 6-1. 구매주문 컬럼 헤더 우클릭
    logger.info('🔍 구매주문 컬럼 헤더 찾는 중...');
    
    // 더 많은 선택자 추가
    const purchaseOrderHeaderSelectors = [
      'div[data-dyn-columnname="NPS_VendPackingSlipSumReportTemp_PurchId"]',
      'div[data-dyn-controlname="NPS_VendPackingSlipSumReportTemp_PurchId"]',
      'div.dyn-headerCell[data-dyn-columnname*="PurchId"]',
      'div.dyn-headerCellLabel[title="구매주문"]',
      '[data-dyn-columnname*="PurchId"]',
      'th:contains("구매주문")',
      'div[title="구매주문"]'
    ];
    
    let headerRightClicked = false;
    
    // JavaScript로 "구매주문" 헤더 찾기 (더 robust한 방법)
    try {
      logger.info('JavaScript로 구매주문 헤더 찾는 중...');
      
      const headerFound = await page.evaluate(() => {
        // 모든 가능한 헤더 요소 검색
        const allHeaders = document.querySelectorAll('th, .dyn-headerCell, [role="columnheader"], div[data-dyn-columnname], div[title]');
        
        for (const header of allHeaders) {
          const text = header.textContent || header.innerText || header.title || '';
          const columnName = header.getAttribute('data-dyn-columnname') || '';
          
          if (text.includes('구매주문') || columnName.includes('PurchId')) {
            // 우클릭 이벤트 발생
            const event = new MouseEvent('contextmenu', {
              bubbles: true,
              cancelable: true,
              button: 2
            });
            header.dispatchEvent(event);
            return true;
          }
        }
        return false;
      });
      
      if (headerFound) {
        logger.info('✅ JavaScript로 구매주문 헤더 우클릭 성공');
        headerRightClicked = true;
      }
    } catch (error) {
      logger.warn(`JavaScript 헤더 우클릭 실패: ${error.message}`);
    }
    
    // 기존 방법으로도 시도
    if (!headerRightClicked) {
      for (const selector of purchaseOrderHeaderSelectors) {
        try {
          logger.info(`구매주문 헤더 선택자 시도: ${selector}`);
          
          if (selector.includes(':contains(')) {
            continue; // CSS :contains()는 지원되지 않으므로 스킵
          }
          
          const headerElement = await page.$(selector);
          if (headerElement) {
            const isVisible = await page.evaluate(el => {
              const style = window.getComputedStyle(el);
              return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
            }, headerElement);
            
            if (isVisible) {
              // 우클릭 실행
              await headerElement.click({ button: 'right' });
              logger.info(`✅ 구매주문 헤더 우클릭 성공: ${selector}`);
              headerRightClicked = true;
              break;
            } else {
              logger.warn(`구매주문 헤더가 보이지 않음: ${selector}`);
            }
          }
        } catch (error) {
          logger.warn(`구매주문 헤더 우클릭 실패: ${selector} - ${error.message}`);
        }
      }
    }
    
    if (!headerRightClicked) {
      logger.error('❌ 구매주문 컬럼 헤더를 찾을 수 없습니다.');
      throw new Error('구매주문 컬럼 헤더를 찾을 수 없습니다.');
    }
    
    // 컨텍스트 메뉴가 나타날 때까지 대기
    logger.info('⏳ 컨텍스트 메뉴 대기 중...');
    await delay(3000);
      // 6-2. "모든 행 내보내기" 메뉴 클릭
    logger.info('🔍 모든 행 내보내기 메뉴 찾는 중...');
    
    let exportMenuClicked = false;
    
    // JavaScript로 "모든 행 내보내기" 메뉴 찾기
    try {
      logger.info('JavaScript로 모든 행 내보내기 메뉴 찾는 중...');
      
      const clicked = await page.evaluate(() => {
        // 1. button-container 내부의 button-label에서 "모든 행 내보내기" 찾기
        const buttonContainers = document.querySelectorAll('.button-container');
        
        for (const container of buttonContainers) {
          const buttonLabel = container.querySelector('.button-label');
          if (buttonLabel) {
            const text = buttonLabel.textContent || buttonLabel.innerText || '';
            if (text.includes('모든 행 내보내기')) {
              // button-container 전체를 클릭
              container.click();
              return { success: true, text: text.trim(), method: 'button-container' };
            }
          }
        }
        
        // 2. 직접 button-label 요소에서 찾기
        const buttonLabels = document.querySelectorAll('.button-label');
        for (const label of buttonLabels) {
          const text = label.textContent || label.innerText || '';
          if (text.includes('모든 행 내보내기')) {
            // 부모 button-container 찾아서 클릭
            const parentContainer = label.closest('.button-container');
            if (parentContainer) {
              parentContainer.click();
              return { success: true, text: text.trim(), method: 'parent-container' };
            } else {
              // 부모가 없으면 label 자체 클릭
              label.click();
              return { success: true, text: text.trim(), method: 'direct-label' };
            }
          }
        }
        
        // 3. 모든 요소에서 텍스트 검색 (기존 방법)
        const allElements = document.querySelectorAll('span, button, [role="button"], [role="menuitem"]');
        
        for (const element of allElements) {
          const text = element.textContent || element.innerText || '';
          if (text.includes('모든 행 내보내기') || text.includes('내보내기') || text.includes('Export')) {
            // 클릭 가능한 부모 요소 찾기
            const clickableParent = element.closest('.button-container, button, [role="button"], [role="menuitem"]') || element;
            clickableParent.click();
            return { success: true, text: text.trim(), method: 'fallback' };
          }
        }
        
        return { success: false };
      });
      
      if (clicked.success) {
        logger.info(`✅ JavaScript로 내보내기 메뉴 클릭 성공 (${clicked.method}): "${clicked.text}"`);
        exportMenuClicked = true;
      }
    } catch (error) {
      logger.warn(`JavaScript 모든 행 내보내기 메뉴 클릭 실패: ${error.message}`);
    }
    
    if (!exportMenuClicked) {
      // 추가 시도: Puppeteer 선택자로 button-container 직접 찾기
      try {
        logger.info('Puppeteer 선택자로 모든 행 내보내기 버튼 찾는 중...');
        
        // button-container 내부에 "모든 행 내보내기" 텍스트가 있는 요소 찾기
        const buttonContainers = await page.$$('.button-container');
        
        for (const container of buttonContainers) {
          try {
            const text = await container.evaluate(el => {
              const label = el.querySelector('.button-label');
              return label ? (label.textContent || label.innerText || '') : '';
            });
            
            if (text.includes('모든 행 내보내기')) {
              await container.click();
              logger.info(`✅ Puppeteer로 내보내기 버튼 클릭 성공: "${text.trim()}"`);
              exportMenuClicked = true;
              break;
            }
          } catch (containerError) {
            logger.warn(`button-container 처리 중 오류: ${containerError.message}`);
          }
        }
      } catch (error) {
        logger.warn(`Puppeteer 모든 행 내보내기 버튼 클릭 실패: ${error.message}`);
      }
    }
    
    if (!exportMenuClicked) {
      logger.error('❌ 모든 행 내보내기 메뉴를 찾을 수 없습니다.');
      throw new Error('모든 행 내보내기 메뉴를 찾을 수 없습니다.');
    }
    
    // 다운로드 대화상자가 나타날 때까지 대기
    logger.info('⏳ 다운로드 대화상자 대기 중...');
    await delay(5000);
    
    // 6-3. "다운로드" 버튼 클릭
    logger.info('🔍 다운로드 버튼 찾는 중...');
    
    let downloadButtonClicked = false;
    
    // JavaScript로 다운로드 버튼 찾기 (더 강력한 로직)
    try {
      logger.info('JavaScript로 다운로드 버튼 찾는 중...');
      
      const clicked = await page.evaluate(() => {
        // 1. "다운로드" 텍스트가 포함된 모든 요소 검색
        const allElements = document.querySelectorAll('button, .button-label, span, [role="button"]');
        
        for (const element of allElements) {
          const text = element.textContent || element.innerText || '';
          if (text.includes('다운로드') || text.includes('Download')) {
            const clickable = element.tagName === 'BUTTON' ? element : element.closest('button, [role="button"], .button-container');
            if (clickable) {
              clickable.click();
              return { success: true, text: text.trim(), method: 'text-search' };
            }
          }
        }
        
        // 2. DownloadButton 관련 속성으로 검색
        const downloadElements = document.querySelectorAll('[name*="DownloadButton"], [id*="DownloadButton"], [data-dyn-controlname*="Download"]');
        for (const el of downloadElements) {
          const button = el.tagName === 'BUTTON' ? el : el.closest('button');
          if (button) {
            button.click();
            return { success: true, method: 'attribute-search' };
          }
        }
        
        // 3. Download 아이콘으로 검색
        const downloadIcons = document.querySelectorAll('.Download-symbol, [class*="download"], [class*="Download"]');
        for (const icon of downloadIcons) {
          const button = icon.closest('button, [role="button"]');
          if (button) {
            button.click();
            return { success: true, method: 'icon-search' };
          }
        }
        
        return { success: false };
      });
      
      if (clicked.success) {
        logger.info(`✅ JavaScript로 다운로드 버튼 클릭 성공 (${clicked.method}): ${clicked.text || 'N/A'}`);
        downloadButtonClicked = true;
      }
    } catch (error) {
      logger.warn(`JavaScript 다운로드 버튼 클릭 실패: ${error.message}`);
    }
    
    if (!downloadButtonClicked) {
      logger.error('❌ 다운로드 버튼을 찾을 수 없습니다.');
      throw new Error('다운로드 버튼을 찾을 수 없습니다.');
    }
    
    // 다운로드 완료 대기
    logger.info('📥 다운로드 실행 중, 완료 대기...');
    await delay(8000);
    
    logger.info('🎉 === 6. 데이터 내보내기 완료 ===');
    
    logger.info('=== 2. 구매 입고내역 조회 페이지 이동 및 데이터 다운로드 완료 ===');
    
    return {
      success: true,
      message: '구매 입고내역 조회 및 데이터 다운로드가 완료되었습니다.'
    };
    
  } catch (error) {
    logger.error(`구매 입고내역 조회 페이지 이동 중 오류: ${error.message}`);
    throw error;
  }
}


// 엑셀 파일에서 특정 셀 값 읽기 함수
function getCellValueFromExcel(filePath, sheetName, cellAddress) {
  try {
    logger.info(`엑셀 파일에서 셀 값 읽기: ${filePath}, 시트: ${sheetName}, 셀: ${cellAddress}`);
    
    const workbook = xlsx.readFile(filePath);
    logger.info(`워크북 로드 완료. 시트 목록: ${Object.keys(workbook.Sheets).join(', ')}`);
    
    // 시트명이 없으면 첫 번째 시트 사용
    const targetSheetName = sheetName || Object.keys(workbook.Sheets)[0];
    const worksheet = workbook.Sheets[targetSheetName];
    
    if (!worksheet) {
      throw new Error(`시트를 찾을 수 없습니다: ${targetSheetName}`);
    }
    
    const cell = worksheet[cellAddress];
    const cellValue = cell ? cell.v : '';
    
    logger.info(`셀 ${cellAddress} 값: "${cellValue}"`);
    return cellValue;
  } catch (error) {
    logger.error(`엑셀 셀 값 읽기 실패: ${error.message}`);
    throw error;
  }
}

// 다운받은 엑셀 파일 경로 찾기 함수 (파일을 열지 않고 경로만 반환)
async function openDownloadedExcel() {
  logger.info('🚀 === 다운받은 엑셀 파일 경로 찾기 시작 ===');
  
  try {
    const os = require('os');
    
    // Windows 기본 다운로드 폴더 경로
    const downloadPath = path.join(os.homedir(), 'Downloads');
    logger.info(`다운로드 폴더 경로: ${downloadPath}`);
    
    // 다운로드 폴더에서 최근 다운받은 엑셀 파일 찾기
    logger.info('최근 다운받은 엑셀 파일 찾는 중...');
    
    const files = fs.readdirSync(downloadPath);
    const excelFiles = files.filter(file => 
      (file.endsWith('.xlsx') || file.endsWith('.xls')) && 
      !file.startsWith('~$') // 임시 파일 제외
    );
    
    if (excelFiles.length === 0) {
      throw new Error('다운로드 폴더에서 엑셀 파일을 찾을 수 없습니다.');
    }
    
    // 파일들을 수정시간 기준으로 정렬하여 가장 최근 파일 찾기
    const excelFilesWithStats = excelFiles.map(file => {
      const filePath = path.join(downloadPath, file);
      const stats = fs.statSync(filePath);
      return {
        name: file,
        path: filePath,
        mtime: stats.mtime
      };
    }).sort((a, b) => b.mtime - a.mtime);
    
    const latestExcelFile = excelFilesWithStats[0];
    logger.info(`최신 엑셀 파일 발견: ${latestExcelFile.name}`);
    logger.info(`파일 경로: ${latestExcelFile.path}`);
    logger.info(`수정시간: ${latestExcelFile.mtime}`);
    
    // 파일이 최근 5분 이내에 다운로드된 것인지 확인
    const fiveMinutesAgo = new Date(Date.now() - 5 * 60 * 1000);
    if (latestExcelFile.mtime < fiveMinutesAgo) {
      logger.warn('⚠️ 발견된 엑셀 파일이 5분 이전에 수정된 파일입니다. 최근 다운로드된 파일이 맞는지 확인하세요.');
    }
    
    // 파일을 열지 않고 경로만 반환
    logger.info('✅ 엑셀 파일 경로를 성공적으로 찾았습니다 (파일을 열지 않음).');
    
    return {
      success: true,
      message: '엑셀 파일 경로를 성공적으로 찾았습니다.',
      filePath: latestExcelFile.path,
      fileName: latestExcelFile.name
    };
    
  } catch (error) {
    logger.error(`엑셀 파일 경로 찾기 중 오류: ${error.message}`);
    
    return {
      success: false,
      error: error.message,
      failedAt: new Date().toISOString(),
      step: '엑셀 파일 경로 찾기'
    };
  }
}

// 3번 RPA 동작: 엑셀 파일 열기 및 매크로 실행 (통합 관리)
async function executeExcelProcessing(page) {
  logger.info('🚀 === 3번 RPA 동작: 엑셀 파일 열기 및 매크로 실행 시작 ===');
  try {
    // 1. 다운로드 폴더에서 최신 엑셀 파일 찾기 (파일을 열지 않고 경로만 획득)
    logger.info('Step 1: 엑셀 파일 경로 찾기 실행 중...');
    const openResult = await openDownloadedExcel();
    if (!openResult.success) {
      throw new Error(openResult.error || '엑셀 파일 경로 찾기에 실패했습니다.');
    }
    logger.info(`✅ Step 1 완료: ${openResult.fileName} (파일을 열지 않고 경로만 획득)`);
    // 2. 매크로 자동 실행 (PowerShell이 엑셀 파일을 열고 매크로 실행)
    logger.info('Step 2: 매크로 자동 실행 시작... (PowerShell이 엑셀 파일을 열고 매크로 실행)');
    const macroResult = await openExcelAndExecuteMacro(openResult.filePath);
    if (!macroResult.success) {
      throw new Error(macroResult.error || '엑셀 매크로 실행에 실패했습니다.');
    }
    logger.info('✅ Step 2 완료: 매크로 실행 성공');
    // 3. 완료 메시지 반환
    logger.info('🎉 === 3번 RPA 동작 완료 ===');
    // 4번 RPA 동작: 대기중인 공급사송장 메뉴 이동 (5초 대기 후 실행)
    logger.info('⏳ 5초 대기 후 4번 RPA 동작(대기중인 공급사송장 메뉴 이동) 시작 예정...');
    await delay(5000);
    
    let step4Status = '4번 RPA 동작 건너뜀';
    if (page) {
      try {
        const pendingResult = await navigateToPendingVendorInvoice(page, openResult.filePath);
        logger.info('4번 RPA 동작 결과:', pendingResult);
        step4Status = '4번 RPA 동작(대기중인 공급사송장 메뉴 이동) 실행 완료';
      } catch (step4Error) {
        logger.error(`4번 RPA 동작 중 오류 발생: ${step4Error.message}`);
        logger.warn('4번 RPA 동작 실패했지만 전체 프로세스는 계속 진행합니다.');
        step4Status = `4번 RPA 동작 실패: ${step4Error.message}`;
      }
    } else {
      logger.warn('4번 RPA 동작을 위한 page 인스턴스가 제공되지 않았습니다.');
    }
    return {
      success: true,
      message: '3번 RPA 동작: 엑셀 파일 매크로 실행이 완료되었습니다.',
      filePath: openResult.filePath,
      fileName: openResult.fileName,
      completedAt: new Date().toISOString(),
      steps: {
        step1: '엑셀 파일 경로 찾기 완료',
        step2: '매크로 실행 완료',
        step3: step4Status
      }
    };
  } catch (error) {
    logger.error(`3번 RPA 동작 중 오류: ${error.message}`);
    return {
      success: false,
      error: error.message,
      failedAt: new Date().toISOString(),
      step: '3번 RPA 동작 (엑셀 파일 열기 및 매크로 실행)'
    };
  }
}

// 4번 RPA 동작: 대기중인 공급사송장 메뉴 이동
async function navigateToPendingVendorInvoice(page, excelFilePath) {
  logger.info('🚀 === 4번 RPA 동작: 대기중인 공급사송장 메뉴 이동 시작 ===');
  try {
    // 1. 검색 버튼 클릭 (2-1과 동일)
    logger.info('검색 버튼 찾는 중...');
    const searchButtonSelectors = [
      '.button-commandRing.Find-symbol',
      'span.Find-symbol',
      '[data-dyn-image-type="Symbol"].Find-symbol',
      '.button-container .Find-symbol'
    ];
    let searchButtonClicked = false;
    for (const selector of searchButtonSelectors) {
      try {
        logger.info(`검색 버튼 선택자 시도: ${selector}`);
        const searchButton = await page.$(selector);
        if (searchButton) {
          const isVisible = await page.evaluate(el => {
            const style = window.getComputedStyle(el);
            return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
          }, searchButton);
          if (isVisible) {
            await searchButton.click();
            logger.info(`검색 버튼 클릭 성공: ${selector}`);
            searchButtonClicked = true;
            break;
          } else {
            logger.warn(`검색 버튼이 보이지 않음: ${selector}`);
          }
        }
      } catch (error) {
        logger.warn(`검색 버튼 클릭 실패: ${selector} - ${error.message}`);
      }
    }
    if (!searchButtonClicked) {
      // JavaScript로 직접 검색 버튼 클릭 시도
      try {
        logger.info('JavaScript로 검색 버튼 직접 클릭 시도...');
        await page.evaluate(() => {
          const searchButtons = document.querySelectorAll('.Find-symbol, [data-dyn-image-type="Symbol"]');
          for (const btn of searchButtons) {
            if (btn.classList.contains('Find-symbol') || btn.getAttribute('data-dyn-image-type') === 'Symbol') {
              btn.click();
              return true;
            }
          }
          return false;
        });
        searchButtonClicked = true;
        logger.info('JavaScript로 검색 버튼 클릭 성공');
      } catch (jsError) {
        logger.error('JavaScript 검색 버튼 클릭 실패:', jsError.message);
      }
    }
    if (!searchButtonClicked) {
      throw new Error('검색 버튼을 찾을 수 없습니다. (4번 RPA)');
    }
    // 검색창이 나타날 때까지 대기
    await delay(2000);
    // 2. "대기중인 공급사송장" 검색어 입력
    logger.info('검색어 입력 중...');
    const searchInputSelectors = [
      'input[type="text"]',
      '.navigationSearchBox input',
      '#NavigationSearchBox',
      'input[placeholder*="검색"]',
      'input[aria-label*="검색"]'
    ];
    let searchInputFound = false;
    const searchTerm = '대기중인 공급사송장';
    for (const selector of searchInputSelectors) {
      try {
        logger.info(`검색 입력창 선택자 시도: ${selector}`);
        await page.waitForSelector(selector, { visible: true, timeout: 5000 });
        await page.click(selector, { clickCount: 3 });
        await page.keyboard.press('Backspace');
        await page.type(selector, searchTerm, { delay: 100 });
        logger.info(`검색어 입력 완료: ${searchTerm}`);
        searchInputFound = true;
        break;
      } catch (error) {
        logger.warn(`검색 입력창 처리 실패: ${selector} - ${error.message}`);
      }
    }
    if (!searchInputFound) {
      throw new Error('검색 입력창을 찾을 수 없습니다. (4번 RPA)');
    }
    // 검색 결과가 나타날 때까지 대기
    await delay(3000);
    // 3. NavigationSearchBox에서 해당 메뉴 클릭
    logger.info('검색 결과에서 대기중인 공급사송장 메뉴 찾는 중...');
    const searchResultSelectors = [
      '.navigationSearchBox',
      '.search-results',
      '.navigation-search-results',
      '[data-dyn-bind*="NavigationSearch"]'
    ];
    let menuClicked = false;
    for (const containerSelector of searchResultSelectors) {
      try {
        const container = await page.$(containerSelector);
        if (container) {
          const menuItems = await page.$$eval(`${containerSelector} *`, (elements) => {
            return elements
              .filter(el => {
                const text = el.textContent || el.innerText || '';
                return text.includes('대기중인 공급사송장');
              })
              .map(el => ({
                text: el.textContent || el.innerText,
                clickable: el.tagName === 'A' || el.tagName === 'BUTTON' || el.onclick || el.getAttribute('role') === 'button'
              }));
          });
          logger.info(`검색 결과 메뉴 항목들:`, menuItems);
          if (menuItems.length > 0) {
            await page.evaluate((containerSel) => {
              const container = document.querySelector(containerSel);
              if (container) {
                const elements = container.querySelectorAll('*');
                for (const el of elements) {
                  const text = el.textContent || el.innerText || '';
                  if (text.includes('대기중인 공급사송장')) {
                    el.click();
                    return true;
                  }
                }
              }
              return false;
            }, containerSelector);
            logger.info('대기중인 공급사송장 메뉴 클릭 완료');
            menuClicked = true;
            break;
          }
        }
      } catch (error) {
        logger.warn(`검색 결과 처리 실패: ${containerSelector} - ${error.message}`);
      }
    }
    if (!menuClicked) {
      // Enter 키로 첫 번째 결과 선택 시도
      logger.info('Enter 키로 검색 결과 선택 시도...');
      await page.keyboard.press('Enter');
      menuClicked = true;
    }
    // 페이지 이동 대기
    logger.info('대기중인 공급사송장 페이지 로딩 대기 중...');
    await delay(5000);
    
    // 4번 RPA 동작 추가 단계들
    logger.info('=== 4번 RPA 동작 추가 단계 시작 ===');
    
    // 4-1. '공급사송장' 탭 클릭
    logger.info('4-1. 공급사송장 탭 찾는 중...');
    try {
      const vendorInvoiceTabClicked = await page.evaluate(() => {
        const spans = document.querySelectorAll('span.appBarTab-headerLabel');
        for (const span of spans) {
          const text = span.textContent || span.innerText || '';
          if (text.includes('공급사송장')) {
            span.click();
            return true;
          }
        }
        return false;
      });
      
      if (vendorInvoiceTabClicked) {
        logger.info('✅ 공급사송장 탭 클릭 성공');
        await delay(3000); // 탭 로딩 대기
      } else {
        logger.warn('⚠️ 공급사송장 탭을 찾을 수 없습니다.');
      }
    } catch (error) {
      logger.warn(`공급사송장 탭 클릭 실패: ${error.message}`);
    }
    
    // 4-2. '제품 입고로 부터' 버튼 클릭
    logger.info('4-2. 제품 입고로 부터 버튼 찾는 중...');
    try {
      const productReceiptButtonClicked = await page.evaluate(() => {
        const buttonContainers = document.querySelectorAll('.button-container');
        for (const container of buttonContainers) {
          const label = container.querySelector('.button-label');
          if (label) {
            const text = label.textContent || label.innerText || '';
            if (text.includes('제품 입고로 부터')) {
              container.click();
              return true;
            }
          }
        }
        return false;
      });
      
      if (productReceiptButtonClicked) {
        logger.info('✅ 제품 입고로 부터 버튼 클릭 성공');
        await delay(3000); // 버튼 클릭 후 로딩 대기
      } else {
        logger.warn('⚠️ 제품 입고로 부터 버튼을 찾을 수 없습니다.');
      }
    } catch (error) {
      logger.warn(`제품 입고로 부터 버튼 클릭 실패: ${error.message}`);
    }
      // 4-3 ~ 4-5. 엑셀 데이터 기반 반복 필터링 처리
    logger.info('4-3 ~ 4-5. 엑셀 데이터 기반 반복 필터링 처리 시작...');
    
    // 먼저 팝업창이 나타날 때까지 대기
    await delay(3000);
    
    try {
      // Step 1: 엑셀에서 A=1이고 B열이 NULL이 아닌 고유한 B값들 수집
      let uniqueBValues = [];
      if (excelFilePath) {
        try {
          logger.info('엑셀에서 A=1이고 B열이 NULL이 아닌 고유한 B값들 수집 중...');
          const workbook = xlsx.readFile(excelFilePath);
          const sheetName = Object.keys(workbook.Sheets)[0]; // 첫 번째 시트
          const worksheet = workbook.Sheets[sheetName];
          
          // 시트 범위 확인
          const range = xlsx.utils.decode_range(worksheet['!ref']);
          const bValues = new Set(); // 중복 제거용
          
          // A=1이고 B열이 NULL이 아닌 행들 찾기
          for (let row = range.s.r + 1; row <= range.e.r; row++) { // 헤더 제외
            const cellA = worksheet[xlsx.utils.encode_cell({ r: row, c: 0 })] || {}; // A열 (0번째 컬럼)
            const cellB = worksheet[xlsx.utils.encode_cell({ r: row, c: 1 })] || {}; // B열 (1번째 컬럼)
            
            const valueA = cellA.v;
            const valueB = cellB.v;
            
            // A=1이고 B가 NULL이 아닌 경우
            // 사이클 넘버 변경
            if (valueA === 5 && valueB && valueB.toString().trim() !== '') {
              bValues.add(valueB.toString().trim());
            }
          }
          
          uniqueBValues = Array.from(bValues);
          logger.info(`수집된 고유한 B값들 (총 ${uniqueBValues.length}개): ${uniqueBValues.join(', ')}`);
        } catch (excelError) {
          logger.warn(`엑셀 데이터 수집 실패: ${excelError.message}`);
          // 백업용 테스트 데이터
          uniqueBValues = ['TEST'];
        }
      } else {
        logger.warn('엑셀 파일 경로가 제공되지 않음, 테스트 데이터 사용');
        uniqueBValues = ['TEST'];
      }
      
      if (uniqueBValues.length === 0) {
        logger.warn('처리할 B값이 없습니다. 기본 테스트 값으로 진행');
        uniqueBValues = ['TEST'];
      }
      
      // Step 2: 각 고유한 B값에 대해 4-3~4-5 순서 반복
      logger.info(`=== ${uniqueBValues.length}개 B값에 대해 순차 처리 시작 ===`);
      
      for (let index = 0; index < uniqueBValues.length; index++) {
        const currentBValue = uniqueBValues[index];
        logger.info(`\n🔄 [${index + 1}/${uniqueBValues.length}] B값 "${currentBValue}" 처리 시작`);
        
        try {
          // 4-3. 구매주문 헤더 클릭
          logger.info(`4-3. 구매주문 헤더 클릭 (B값: "${currentBValue}")`);
          
          const purchaseOrderHeaderClicked = await page.evaluate(() => {
            const dialogPopup = document.querySelector('.dialog-popup-content');
            if (!dialogPopup) {
              return { success: false, error: '팝업창을 찾을 수 없습니다.' };
            }
            
            // 구매주문 헤더 찾기
            const popupHeaders = dialogPopup.querySelectorAll('.dyn-headerCellLabel._11w1prk, .dyn-headerCellLabel');
            for (const header of popupHeaders) {
              const title = (header.getAttribute('title') || '').trim();
              const text = (header.textContent || header.innerText || '').trim();
              
              if (title === '구매주문' || text === '구매주문') {
                header.click();
                return { 
                  success: true, 
                  method: 'popup-header-text', 
                  title: title, 
                  text: text
                };
              }
            }
            
            // 백업: PurchOrder 포함된 요소 찾기
            const purchaseOrderElements = dialogPopup.querySelectorAll('[data-dyn-columnname*="PurchOrder"], [data-dyn-controlname*="PurchOrder"]');
            for (const element of purchaseOrderElements) {
              const headerLabel = element.querySelector('.dyn-headerCellLabel._11w1prk') || 
                                element.querySelector('.dyn-headerCellLabel');
              if (headerLabel) {
                headerLabel.click();
                return { 
                  success: true, 
                  method: 'popup-columnname-partial'
                };
              }
            }
            
            return { success: false, error: '팝업창 내에서 구매주문 헤더를 찾을 수 없습니다.' };
          });
          
          if (!purchaseOrderHeaderClicked.success) {
            logger.warn(`⚠️ 구매주문 헤더 클릭 실패 (B값: "${currentBValue}"): ${purchaseOrderHeaderClicked.error}`);
            continue; // 다음 B값으로 넘어감
          }
          
          logger.info(`✅ 구매주문 헤더 클릭 성공 (${purchaseOrderHeaderClicked.method})`);
          await delay(1000); // 헤더 클릭 후 필터창 로딩 대기
          
          // 4-4. 필터 입력창에 현재 B값 입력
          logger.info(`4-4. 필터 입력창에 B값 "${currentBValue}" 입력 중...`);
          
          const filterInputResult = await page.evaluate((value) => {
            const filterPopup = document.querySelector('.columnHeader-popup');
            if (!filterPopup) return { success: false, error: '필터 팝업창이 존재하지 않음' };
            
            const inputSelectors = [
              'input[role="combobox"]',
              'input.textbox.field',
              'input[type="text"]',
              'input[name*="Filter"]'
            ];
            
            for (const selector of inputSelectors) {
              const input = filterPopup.querySelector(selector);
              if (input && input.offsetParent !== null) {
                // 기존 값 클리어
                input.focus();
                input.value = '';
                input.dispatchEvent(new Event('input', { bubbles: true }));
                
                // 새 값 입력
                input.value = value;
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
                
                return { 
                  success: true, 
                  method: 'direct-input',
                  selector: selector,
                  value: value
                };
              }
            }
            
            return { success: false, error: '필터 입력창을 찾을 수 없습니다.' };
          }, currentBValue);
          
          if (!filterInputResult.success) {
            logger.warn(`⚠️ 필터 입력 실패 (B값: "${currentBValue}"): ${filterInputResult.error}`);
            continue; // 다음 B값으로 넘어감
          }
          
          logger.info(`✅ 필터 입력 성공: "${filterInputResult.value}"`);
          
          // 4-5. Enter 키로 필터 적용
          logger.info('4-5. Enter 키로 필터 적용 중...');
          await delay(500);
          await page.keyboard.press('Enter');
          logger.info('✅ Enter 키로 필터 적용 완료');
          
          // 필터링 완료 대기 (단축: 10초 → 5초)
          logger.info('필터링 완료 대기 중... (5초)');
          await delay(5000);
          
          // 4-5-2. All Check 버튼 클릭
          logger.info('4-5-2. All Check 버튼 클릭 중...');
          
          const allCheckClicked = await page.evaluate(() => {
            // All Check 버튼 찾기
            const allCheckSpan = document.querySelector('#PurchJournalSelect_PackingSlip_45_NPS_AllCheck_label');
            if (allCheckSpan && allCheckSpan.textContent.trim() === 'All Check') {
              allCheckSpan.click();
              return { 
                success: true, 
                method: 'exact-span-id-AllCheck',
                text: allCheckSpan.textContent.trim()
              };
            }
            
            // 백업: span.button-label에서 "All Check" 찾기
            const allSpans = document.querySelectorAll('span.button-label');
            for (const span of allSpans) {
              const spanText = (span.textContent || span.innerText || '').trim();
              if (spanText === 'All Check') {
                span.click();
                return { 
                  success: true, 
                  method: 'span-text-AllCheck',
                  text: spanText
                };
              }
            }
            
            return { success: false, error: 'All Check 버튼을 찾을 수 없습니다.' };
          });
          
          if (allCheckClicked.success) {
            logger.info(`✅ All Check 버튼 클릭 성공 (${allCheckClicked.method}): "${allCheckClicked.text}"`);
            await delay(1000); // All Check 처리 대기
          } else {
            logger.warn(`⚠️ All Check 버튼 클릭 실패 (B값: "${currentBValue}"): ${allCheckClicked.error}`);
          }
          
          logger.info(`🎉 [${index + 1}/${uniqueBValues.length}] B값 "${currentBValue}" 처리 완료\n`);
          
          // 사이클 완료 후 Alt + Enter 입력
          logger.info('사이클 완료 후 Alt + Enter 입력 중...');
          await page.keyboard.down('Alt');
          await page.keyboard.press('Enter');
          await page.keyboard.up('Alt');
          logger.info('✅ Alt + Enter 입력 완료');
          
          // 다음 B값 처리를 위한 짧은 대기 (1초)
          if (index < uniqueBValues.length - 1) {
            await delay(1000);
          }
          
        } catch (currentBError) {
          logger.warn(`❌ B값 "${currentBValue}" 처리 중 오류: ${currentBError.message}`);
          continue; // 다음 B값으로 넘어감
        }
      }
      
      logger.info(`🎉 === 모든 B값 처리 완료 (총 ${uniqueBValues.length}개) ===`);
      
    } catch (error) {
      logger.warn(`반복 필터링 처리 중 오류: ${error.message}`);
    }
    
    // 4-6. 프로세스 완료
    logger.info('=== 4번 RPA 동작: 대기중인 공급사송장 메뉴 이동 및 엑셀 데이터 처리 완료 ===');

    // ==== 추가 기능 시작 ====  
    logger.info('추가기능: “송장일” 셀 옆 캘린더 아이콘(svg._1dciz1s) 더블클릭 → TEST 입력');

    try {
      // 1) 캘린더 아이콘(svg)을 기다림
      const iconSelector = 'svg._1dciz1s';
      await page.waitForSelector(iconSelector, { visible: true, timeout: 10000 });

      // 2) 첫 번째 아이콘을 가져와 더블클릭
      const icons = await page.$$(iconSelector);
      if (icons.length === 0) {
        throw new Error('캘린더 아이콘을 찾을 수 없습니다.');
      }
      const targetIcon = icons[0];
      await targetIcon.click({ clickCount: 2, delay: 100 });
      logger.info('✅ 캘린더 아이콘 더블클릭 완료');

      // 3) 아이콘 클릭으로 열리는 팝업 내 입력창(Combobox) 선택
      //    – 팝업이 input[aria-label="송장일"] 이 보일 때까지 대기
      const inputSelector = 'input[aria-label="송장일"]';
      await page.waitForSelector(inputSelector, { visible: true, timeout: 5000 });

      // 4) 전체 선택 보강(3번 클릭) 후 기존 값 삭제
      await page.click(inputSelector, { clickCount: 3, delay: 50 });
      await page.keyboard.press('Backspace');

      // 5) "TEST" 입력
      await page.type(inputSelector, 'TEST', { delay: 100 });
      logger.info('✅ "TEST" 입력 완료');
    } catch (err) {
      logger.warn('⚠️ 추가 기능 실패: ' + err.message);
    }
    
    // ==== 추가 기능 끝 ====

    return { success: true, message: '4번 RPA 동작: 대기중인 공급사송장 메뉴 이동 및 엑셀 데이터 처리 완료' };
  } catch (error) {
    logger.error(`4번 RPA 동작 중 오류: ${error.message}`);
    return { success: false, error: error.message, step: '4번 RPA 동작 (대기중인 공급사송장 메뉴 이동)' };
  }
}

// 엑셀 파일 열기 및 매크로 자동 실행 함수
async function openExcelAndExecuteMacro(excelFilePath) {
  const { exec } = require('child_process');
  const { promisify } = require('util');
  const os = require('os');
  const execAsync = promisify(exec);
  
  logger.info('🚀 === 엑셀 파일 열기 및 매크로 자동 실행 시작 ===');
  logger.info(`대상 엑셀 파일: ${excelFilePath}`);
  
  try {
    // VBA 코드 정의
    const vbaCode = `
Sub GroupBy_I_Z_And_Process()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, groupNum As Long
    Dim key As String
    Dim groupMap As Object, groupSums As Object, groupDesc As Object
    Dim maturityDate As Date, adjustedSum As Double
    Dim maturityCol As Long, descCol As Long, taxDateCol As Long
    Dim gDate As Variant
    Dim currentGroup As Long, nextGroup As Long
    Dim lastCol As Long
    Dim pText As String, jText As String

    Set ws = ActiveSheet ' Use active sheet instead of specific name
    Set groupMap = CreateObject("Scripting.Dictionary")
    Set groupSums = CreateObject("Scripting.Dictionary")
    Set groupDesc = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False   ' Turn off screen update
    Application.Calculation = xlCalculationManual   ' Turn off auto calc

    ' Find last row in column I
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

    ' Sort by I and Z columns
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("I2:I" & lastRow), Order:=xlAscending
    ws.Sort.SortFields.Add Key:=ws.Range("Z2:Z" & lastRow), Order:=xlAscending
    With ws.Sort
        .SetRange ws.Range("A1:AG" & lastRow)
        .Header = xlYes
        .Apply
    End With

    ' Insert Group Number column at A
    ws.Columns("A").Insert Shift:=xlToRight
    ws.Cells(1, 1).Value = "Group Number"

    groupNum = 1

    ' Assign group number, sum AG, make invoice description
    For i = 2 To lastRow
        key = ws.Cells(i, "I").Value & "|" & ws.Cells(i, "Z").Value

        If Not groupMap.exists(key) Then
            groupMap(key) = groupNum
            groupSums(groupNum) = 0

            ' Invoice description: Month(gDate) & " month " & P & "_" & J
            gDate = ws.Cells(i, "G").Value
            pText = ws.Cells(i, "P").Value
            jText = ws.Cells(i, "J").Value

            If IsDate(gDate) Then
                groupDesc(groupNum) = Month(gDate) & " month " & pText & "_" & jText
            Else
                groupDesc(groupNum) = "Date Error " & pText & "_" & jText
            End If

            groupNum = groupNum + 1
        End If

        ws.Cells(i, 1).Value = groupMap(key)
        groupSums(groupMap(key)) = groupSums(groupMap(key)) + Val(ws.Cells(i, "AG").Value)
    Next i

    ' Add columns: Maturity Date, Invoice Description, Tax Invoice Date
    maturityCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    ws.Cells(1, maturityCol).Value = "Maturity Date"

    descCol = maturityCol + 1
    ws.Cells(1, descCol).Value = "Invoice Description"

    taxDateCol = descCol + 1
    ws.Cells(1, taxDateCol).Value = "Tax Invoice Date"

    ' Fill Maturity Date, Invoice Description, Tax Invoice Date
    For i = 2 To lastRow
        Dim gNum As Long
        gNum = ws.Cells(i, 1).Value
        gDate = ws.Cells(i, "G").Value

        If IsDate(gDate) Then
            adjustedSum = groupSums(gNum) * 1.1
            If adjustedSum < 10000000 Then
                maturityDate = WorksheetFunction.EoMonth(gDate, 1) ' End of next month
            Else
                maturityDate = WorksheetFunction.EoMonth(gDate, 2) ' End of following month
            End If
            ws.Cells(i, maturityCol).Value = maturityDate

            ' Tax Invoice Date: end of the month for G column
            ws.Cells(i, taxDateCol).Value = WorksheetFunction.EoMonth(gDate, 0)
        Else
            ws.Cells(i, maturityCol).Value = "Date Error"
            ws.Cells(i, taxDateCol).Value = "Date Error"
        End If

        ' Enter Invoice Description
        ws.Cells(i, descCol).Value = groupDesc(gNum)
    Next i

    ' Apply date format to Tax Invoice Date column
    ws.Range(ws.Cells(2, taxDateCol), ws.Cells(lastRow, taxDateCol)).NumberFormat = "yyyy-mm-dd"

    ' Add line between groups
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 2 To lastRow - 1
        currentGroup = ws.Cells(i, 1).Value
        nextGroup = ws.Cells(i + 1, 1).Value

        If currentGroup <> nextGroup Then
            With ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
        End If
    Next i

    Application.ScreenUpdating = True    ' Turn on screen update
    Application.Calculation = xlCalculationAutomatic   ' Turn on auto calc

End Sub
`;

    // 임시 PowerShell 스크립트 생성
    const tempDir = os.tmpdir();
    const psScriptPath = path.join(tempDir, `excel_macro_${Date.now()}.ps1`);
    
    // PowerShell 스크립트 내용 (VBA 코드를 직접 포함)
    const psScript = `
# Excel 매크로 자동 실행 PowerShell 스크립트
param(
    [string]$ExcelFilePath = "${excelFilePath.replace(/\\/g, '\\\\')}"
)

Write-Host "Excel 매크로 자동 실행 스크립트 시작"
Write-Host "대상 파일: $ExcelFilePath"

try {
    # COM 객체 생성
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    Write-Host "Excel 애플리케이션 생성 완료"
    
    # 기존에 열린 워크북이 있는지 확인
    $workbook = $null
    $fileName = Split-Path $ExcelFilePath -Leaf
    
    foreach ($wb in $excel.Workbooks) {
        if ($wb.Name -eq $fileName) {
            $workbook = $wb
            Write-Host "기존에 열린 워크북 사용: $fileName"
            break
        }
    }
    
    # 워크북이 없으면 새로 열기
    if ($workbook -eq $null) {
        if (Test-Path $ExcelFilePath) {
            $workbook = $excel.Workbooks.Open($ExcelFilePath)
            Write-Host "워크북 열기 완료: $ExcelFilePath"
        } else {
            throw "파일을 찾을 수 없습니다: $ExcelFilePath"
        }
    }
    
    # 워크시트 선택
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Activate()
    
    Write-Host "워크시트 활성화 완료"
    
    # 기존 VBA 모듈 제거
    $vbaProject = $workbook.VBProject
    for ($i = $vbaProject.VBComponents.Count; $i -ge 1; $i--) {
        $component = $vbaProject.VBComponents.Item($i)
        if ($component.Type -eq 1) {  # vbext_ct_StdModule
            $vbaProject.VBComponents.Remove($component)
            Write-Host "기존 VBA 모듈 제거: $($component.Name)"
        }
    }
    
    # 새 VBA 모듈 추가
    $vbaModule = $vbaProject.VBComponents.Add(1)  # vbext_ct_StdModule
    $vbaModule.Name = "GroupProcessModule"
    
    Write-Host "새 VBA 모듈 추가 완료"
    
    # VBA 코드 추가 - 잠시 대기 후 추가
    Start-Sleep -Milliseconds 500
    
    # VBA 코드 추가
    $vbaCode = @"
${vbaCode}
"@;
    
    $vbaModule.CodeModule.AddFromString($vbaCode)
    Write-Host "VBA 코드 추가 완료"
    
    # 매크로 실행 전 대기
    Start-Sleep -Seconds 2
   
    Write-Host "VBA 프로젝트 준비 완료, 매크로 실행 중..."
    
    # 매크로 실행 - 정확한 함수명 사용
    try {
        $excel.Run("GroupBy_I_Z_And_Process")
        Write-Host "매크로 실행 완료"
    } catch {
        Write-Host "매크로 실행 실패: $($_.Exception.Message)"
        # 대안으로 모듈명.함수명 형태로 시도
        try {
            $excel.Run("GroupProcessModule.GroupBy_I_Z_And_Process")
            Write-Host "모듈명 포함 매크로 실행 완료"
        } catch {
            Write-Host "모듈명 포함 매크로 실행도 실패: $($_.Exception.Message)"
            throw "매크로 실행에 실패했습니다."
        }
    }
    
    # 매크로 실행 후 파일 저장
    Start-Sleep -Seconds 2
    Write-Host "매크로 실행 후 파일 저장 중..."
    
    try {
        $workbook.Save()
        Write-Host "파일 저장 완료"
    } catch {
        Write-Host "파일 저장 실패: $($_.Exception.Message)"
        # 다른 이름으로 저장 시도
        try {
            $savePath = $ExcelFilePath -replace '\.xlsx$', '_processed.xlsx'
            $workbook.SaveAs($savePath)
            Write-Host "다른 이름으로 저장 완료: $savePath"
        } catch {
            Write-Host "다른 이름으로 저장도 실패: $($_.Exception.Message)"
            throw "파일 저장에 실패했습니다."
        }
    }
    
    # Excel을 보이게 설정
    $excel.Visible = $true
    $excel.DisplayAlerts = $true
    
    Write-Host "Excel 매크로 자동 실행 완료"
    
} catch {
    Write-Host "오류 발생: $($_.Exception.Message)"
    if ($excel) {
        $excel.Visible = $true
        $excel.DisplayAlerts = $true
    }
    exit 1
}
`;

    // PowerShell 스크립트 파일 저장
    fs.writeFileSync(psScriptPath, psScript, 'utf8');
    logger.info(`PowerShell 스크립트 생성 완료: ${psScriptPath}`);
    
    // PowerShell 스크립트 실행
    logger.info('PowerShell 스크립트 실행 중...');
    const result = await execAsync(`powershell -ExecutionPolicy Bypass -File "${psScriptPath}"`, {
      timeout: 60000, // 60초 타임아웃
      encoding: 'utf8'
    });
    
    if (result.stdout) {
      logger.info('PowerShell 실행 결과:');
      logger.info(result.stdout);
    }
    
    if (result.stderr) {
      logger.warn('PowerShell 실행 경고:');
      logger.warn(result.stderr);
    }
    
    // 임시 파일 정리
    try {
      fs.unlinkSync(psScriptPath);
      logger.info('임시 PowerShell 스크립트 파일 정리 완료');
    } catch (cleanupError) {
      logger.warn(`임시 파일 정리 실패: ${cleanupError.message}`);
    }
    
    logger.info('✅ 엑셀 매크로 자동 실행 완료');
    
    return {
      success: true,
      message: '엑셀 매크로가 성공적으로 실행되었습니다.',
      filePath: excelFilePath
    };
    
  } catch (error) {
    logger.error(`엑셀 매크로 실행 중 오류: ${error.message}`);
    
    return {
      success: false,
      error: error.message,
      failedAt: new Date().toISOString(),
      step: '엑셀 매크로 실행'
    };
  }
}

// 모듈 내보내기
module.exports = {
  setCredentials,
  getCredentials,
  connectToD365,
  waitForDataTable,
  processInvoice: connectToD365, // 전체 프로세스 기능 활성화 (connectToD365와 동일한 함수)
  openDownloadedExcel,
  openExcelAndExecuteMacro,
  executeExcelProcessing, // 3번 동작: 엑셀 파일 열기 및 매크로 실행 통합 관리
  navigateToPendingVendorInvoice, // 4번 동작: 대기중인 공급사송장 메뉴 이동
};
