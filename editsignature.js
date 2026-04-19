  // editsignature.js - Simplified signature manager                                                                                                       
                                                                                                                                                           
  Office.onReady(function() {                                                                                                                              
      console.log("Office.js ready - loading edit signature page");                                                                                        
      loadUserInfo();                                                                                                                                      
                                                                                                                                                           
      // Hidden keyboard shortcut: Ctrl+Shift+D to toggle Clear Cache button                                                                               
      document.addEventListener('keydown', function(e) {
          if (e.ctrlKey && e.shiftKey && e.key === 'D') {                                                                                                  
              e.preventDefault();                                                                                                                          
              const btn = document.getElementById('clearCacheButton');                                                                                     
              if (btn) {                                                                                                                                   
                  btn.style.display = btn.style.display === 'none' ? 'inline-flex' : 'none';                                                               
                  console.log('Debug mode:', btn.style.display === 'none' ? 'OFF' : 'ON');                                                                 
              }                                                                                                                                            
          }                                                                                                                                                
      });                                                                                                                                                  
  });                                                                                                                                                      
   
  /**                                                                                                                                                      
   * Load user info from /signature endpoint (one call gets everything)
   */                                                                                                                                                      
  async function loadUserInfo() {
      try {                                                                                                                                                
          const userEmail = Office.context.mailbox.userProfile.emailAddress;
          const API_BASE_URL = window.location.hostname === 'localhost'                                                                                    
              ? 'http://localhost:3001'
              : 'https://lilly-signature-addin.dc.lilly.com';                                                                                              
                                                                                                                                                           
          const response = await fetch(`${API_BASE_URL}/signature?email=${encodeURIComponent(userEmail)}`, {                                               
              credentials: 'include',                                                                                                                      
              mode: 'cors'                                                                                                                                 
          });                                                                                                                                              
                                                                                                                                                           
          if (!response.ok) throw new Error('Failed to fetch');                                                                                            
                  
          const data = await response.json();                                                                                                              
                  
          // Display the data                                                                                                                              
          populateField('display_name_readonly', data.displayName);
          populateField('email_readonly', data.email);                                                                                                     
          populateField('department_readonly', data.department);                                                                                           
          populateField('office_phone_readonly', data.officePhone, 'officePhoneAction');                                                                   
          populateField('mobile_readonly', data.mobilePhone, 'mobilePhoneAction');                                                                         
          populateField('location_readonly', data.country, 'locationAction');                                                                              
                                                                                                                                                           
          // Load title — check for existing override first, then fall back to Workday value                                                               
          await loadTitleField(userEmail, data.jobTitle);                                                                                                  
                                                                                                                                                           
          console.log('✓ User info loaded successfully');                                                                                                  
                                                                                                                                                           
      } catch (error) {                                                                                                                                    
          console.error("Failed to load user info:", error);
          showError("Could not load your information. Please try refreshing the page.");                                                                   
      }
  }                                                                                                                                                        
                  
  /**                                                                                                                                                      
   * Populate a single field
   */
  function populateField(elementId, value, actionId) {
      const element = document.getElementById(elementId);                                                                                                  
      if (!element) return;                                                                                                                                
                                                                                                                                                           
      if (value && value.trim() !== "") {                                                                                                                  
          element.textContent = value;
          element.classList.remove('empty');                                                                                                               
      } else {                                                                                                                                             
          element.innerHTML = '<span style="color: #9ca3af; font-style: italic;">Not provided</span>';
          element.classList.add('empty');                                                                                                                  
                                                                                                                                                           
          // Show "Update in Workday" link if field has an action                                                                                          
          if (actionId) {                                                                                                                                  
              const actionElement = document.getElementById(actionId);                                                                                     
              if (actionElement) actionElement.style.display = 'block';                                                                                    
          }                                                                                                                                                
      }                                                                                                                                                    
  }                                                                                                                                                        
                  
  /**                                                                                                                                                      
   * Copy signature from database
   */
  async function copySignatureFromDatabase() {
      const copyBtn = document.getElementById('copyButton');
      const successMsg = document.getElementById('copySuccessMessage');                                                                                    
      const originalHTML = copyBtn.innerHTML;                                                                                                              
                                                                                                                                                           
      copyBtn.disabled = true;                                                                                                                             
      copyBtn.innerHTML = `                                                                                                                                
          <svg class="btn-icon btn-icon-spinner" viewBox="0 0 20 20" fill="currentColor">                                                                  
              <path fill-rule="evenodd" d="M4 2a1 1 0 011 1v2.101a7.002 7.002 0 0111.601 2.566 1 1 0 11-1.885.666A5.002 5.002 0 005.999 7H9a1 1 0 010 2H4a1
   1 0 01-1-1V3a1 1 0 011-1zm.008 9.057a1 1 0 011.276.61A5.002 5.002 0 0014.001 13H11a1 1 0 110-2h5a1 1 0 011 1v5a1 1 0 11-2 0v-2.101a7.002 7.002 0        
  01-11.601-2.566 1 1 0 01.61-1.276z" clip-rule="evenodd"/>                                                                                                
          </svg>                                                                                                                                           
          <span class="btn-text">Loading...</span>
      `;                                                                                                                                                   
                                                                                                                                                           
      try {                                                                                                                                                
          const userEmail = Office.context.mailbox.userProfile.emailAddress;                                                                               
          const API_BASE_URL = window.location.hostname === 'localhost'                                                                                    
              ? 'http://localhost:3001'
              : 'https://lilly-signature-addin.dc.lilly.com';                                                                                              
                                                                                                                                                           
          const response = await fetch(`${API_BASE_URL}/signature?email=${encodeURIComponent(userEmail)}`);                                                
          if (!response.ok) throw new Error('Failed to fetch signature');                                                                                  
                                                                                                                                                           
          const data = await response.json();
                                                                                                                                                           
          // Copy HTML to clipboard using execCommand fallback (Clipboard API is blocked in Office Add-ins)                                                
          const tempDiv = document.createElement('div');
          tempDiv.contentEditable = true;                                                                                                                  
          tempDiv.innerHTML = data.signatureHTML;                                                                                                          
          tempDiv.style.position = 'fixed';                                                                                                                
          tempDiv.style.left = '-9999px';                                                                                                                  
          document.body.appendChild(tempDiv);                                                                                                              
                                                                                                                                                           
          // Select the content                                                                                                                            
          const range = document.createRange();                                                                                                            
          range.selectNodeContents(tempDiv);                                                                                                               
          const selection = window.getSelection();
          selection.removeAllRanges();                                                                                                                     
          selection.addRange(range);                                                                                                                       
                                                                                                                                                           
          // Copy                                                                                                                                          
          const success = document.execCommand('copy');
                                                                                                                                                           
          // Cleanup
          selection.removeAllRanges();                                                                                                                     
          document.body.removeChild(tempDiv);                                                                                                              
                                                                                                                                                           
          if (!success) {                                                                                                                                  
              throw new Error('execCommand copy failed');                                                                                                  
          }                                                                                                                                                
   
          // Show success                                                                                                                                  
          successMsg.style.display = 'flex';
          copyBtn.innerHTML = `                                                                                                                            
              <svg class="btn-icon" viewBox="0 0 20 20" fill="currentColor">
                  <path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0     
  011.414 0z" clip-rule="evenodd"/>                                                                                                                        
              </svg>                                                                                                                                       
              <span class="btn-text">Copied!</span>                                                                                                        
          `;                                                                                                                                               
   
          setTimeout(() => {                                                                                                                               
              successMsg.style.display = 'none';
              copyBtn.innerHTML = originalHTML;
              copyBtn.disabled = false;
          }, 3000);
                                                                                                                                                           
      } catch (error) {
          console.error('Failed to copy:', error);                                                                                                         
          // Show error in UI instead of alert (alert is not supported in Office Add-ins)                                                                  
          successMsg.style.display = 'flex';                                                                                                               
          successMsg.style.background = '#fee2e2';                                                                                                         
          successMsg.style.color = '#991b1b';                                                                                                              
          successMsg.innerHTML = `                                                                                                                         
              <svg class="tip-icon" viewBox="0 0 20 20" fill="currentColor">                                                                               
                  <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414      
  1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clip-rule="evenodd"/>                  
              </svg>                                                                                                                                       
              <span>Failed to copy. Please try again.</span>                                                                                               
          `;                                                                                                                                               
          copyBtn.innerHTML = originalHTML;
          copyBtn.disabled = false;                                                                                                                        
                  
          setTimeout(() => {                                                                                                                               
              successMsg.style.display = 'none';
              successMsg.style.background = '#d1fae5';                                                                                                     
              successMsg.style.color = '#065f46';                                                                                                          
          }, 3000);                                                                                                                                        
      }                                                                                                                                                    
  }                                                                                                                                                        
                  
  /**
   * Clear all caches and reload data
   */                                                                                                                                                      
  function clearCache() {
      console.log("🗑️  Clearing all caches...");                                                                                                            
                  
      const clearButton = document.getElementById('clearCacheButton');                                                                                     
      const originalHTML = clearButton.innerHTML;
                                                                                                                                                           
      clearButton.disabled = true;
      clearButton.innerHTML = `
          <svg class="btn-icon btn-icon-spinner" viewBox="0 0 20 20" fill="currentColor">                                                                  
              <path fill-rule="evenodd" d="M4 2a1 1 0 011 1v2.101a7.002 7.002 0 0111.601 2.566 1 1 0 11-1.885.666A5.002 5.002 0 005.999 7H9a1 1 0 010 2H4a1
   1 0 01-1-1V3a1 1 0 011-1zm.008 9.057a1 1 0 011.276.61A5.002 5.002 0 0014.001 13H11a1 1 0 110-2h5a1 1 0 011 1v5a1 1 0 11-2 0v-2.101a7.002 7.002 0        
  01-11.601-2.566 1 1 0 01.61-1.276z" clip-rule="evenodd"/>                                                                                                
          </svg>                                                                                                                                           
          <span class="btn-text">Clearing...</span>
      `;
                                                                                                                                                           
      try {
          // Clear localStorage                                                                                                                            
          localStorage.removeItem('lilly_user_info');
                                                                                                                                                           
          // Clear sessionStorage
          if (typeof sessionStorage !== 'undefined') {                                                                                                     
              sessionStorage.removeItem('user_info_session_cache');                                                                                        
          }                                                                                                                                                
                                                                                                                                                           
          // Clear roamingSettings                                                                                                                         
          Office.context.roamingSettings.remove('user_info_cache');
          Office.context.roamingSettings.remove('user_info_timestamp');                                                                                    
          Office.context.roamingSettings.remove('lilly_user_info');                                                                                        
          Office.context.roamingSettings.remove('user_info');                                                                                              
                                                                                                                                                           
          Office.context.roamingSettings.saveAsync(function(result) {                                                                                      
              if (result.status === Office.AsyncResultStatus.Succeeded) {                                                                                  
                  console.log("✓ Cache cleared");                                                                                                          
                  clearButton.innerHTML = `                                                                                                                
                      <svg class="btn-icon" viewBox="0 0 20 20" fill="currentColor">                                                                       
                          <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0       
  00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd"/>                                                                                           
                      </svg>                                                                                                                               
                      <span class="btn-text">Cache Cleared!</span>                                                                                         
                  `;                                                                                                                                       
                                                                                                                                                           
                  setTimeout(() => location.reload(), 1000);                                                                                               
              } else {
                  console.error("Failed to clear cache:", result.error);                                                                                   
                  clearButton.disabled = false;                                                                                                            
                  clearButton.innerHTML = originalHTML;                                                                                                    
                  alert('Failed to clear cache. Please try again.');                                                                                       
              }                                                                                                                                            
          });
      } catch (error) {                                                                                                                                    
          console.error("Error clearing cache:", error);
          clearButton.disabled = false;                                                                                                                    
          clearButton.innerHTML = originalHTML;                                                                                                            
          alert('Error clearing cache: ' + error.message);                                                                                                 
      }                                                                                                                                                    
  }                                                                                                                                                        
   
  function showError(message) {                                                                                                                            
      alert(message);
  }                                                                                                                                                        
                                                                                                                                                           
  /**                                                                                                                                                      
   * Load title field — show existing override if present, else show Workday value                                                                         
   */                                                                                                                                                      
  async function loadTitleField(userEmail, workdayTitle) {
      const input = document.getElementById('job_title_input');                                                                                            
      if (!input) return;                                                                                                                                  
                                                                                                                                                           
      try {                                                                                                                                                
          const API_BASE_URL = window.location.hostname === 'localhost' ? 'http://localhost:3001' : 'https://lilly-signature-addin.dc.lilly.com';          
          const response = await fetch(`${API_BASE_URL}/api/user/title-override?email=${encodeURIComponent(userEmail)}`);                                  
          if (response.ok) {                                                                                                                               
              const data = await response.json();                                                                                                          
              input.value = data.customTitle || workdayTitle || '';                                                                                        
              if (data.customTitle) {                                                                                                                      
                  input.dataset.hasOverride = 'true';                                                                                                      
              }                                                                                                                                            
          } else {                                                                                                                                         
              input.value = workdayTitle || '';                                                                                                            
          }                                                                                                                                                
      } catch {
          input.value = workdayTitle || '';                                                                                                                
      }           
  }                                                                                                                                                        
   
  /**                                                                                                                                                      
   * Save the user's custom title override
   */                                                                                                                                                      
  async function saveTitle() {
      const input = document.getElementById('job_title_input');                                                                                            
      const saveBtn = document.getElementById('saveTitleBtn');
      const msg = document.getElementById('titleSaveMsg');                                                                                                 
      if (!input || !saveBtn || !msg) return;                                                                                                              
                                                                                                                                                           
      const newTitle = input.value.trim();                                                                                                                 
      if (!newTitle) {                                                                                                                                     
          msg.textContent = 'Title cannot be empty.';                                                                                                      
          msg.style.color = '#b91c1c';                                                                                                                     
          msg.style.display = 'block';                                                                                                                     
          return;                                                                                                                                          
      }                                                                                                                                                    
                                                                                                                                                           
      saveBtn.disabled = true;                                                                                                                             
      saveBtn.textContent = 'Saving...';
                                                                                                                                                           
      try {       
          const userEmail = Office.context.mailbox.userProfile.emailAddress;                                                                               
          const API_BASE_URL = window.location.hostname === 'localhost' ? 'http://localhost:3001' : 'https://lilly-signature-addin.dc.lilly.com';          
                                                                                                                                                           
          const response = await fetch(`${API_BASE_URL}/api/user/title-override`, {                                                                        
              method: 'POST',                                                                                                                              
              headers: { 'Content-Type': 'application/json' },                                                                                             
              credentials: 'include',                                                                                                                      
              body: JSON.stringify({ email: userEmail, customTitle: newTitle })
          });                                                                                                                                              
                                                                                                                                                           
          if (!response.ok) throw new Error('Server error');                                                                                               
                                                                                                                                                           
          msg.textContent = 'Title saved. Your signature will update within 3 days.';                                                                      
          msg.style.color = '#065f46';
          msg.style.display = 'block';                                                                                                                     
          setTimeout(() => { msg.style.display = 'none'; }, 4000);                                                                                         
      } catch {                                                                                                                                            
          msg.textContent = 'Failed to save. Please try again.';                                                                                           
          msg.style.color = '#b91c1c';                                                                                                                     
          msg.style.display = 'block';
      } finally {                                                                                                                                          
          saveBtn.disabled = false;
          saveBtn.textContent = 'Save';                                                                                                                    
      }                                                                                                                                                    
  }   
