        lucide.createIcons();
        const API_BASE = window.location.origin;

        // Logging System Setup
        const systemLogs = [];
        const logContentEl = document.getElementById('logContent');

        function addLog(msg, type = 'INFO') {
            const timestamp = new Date().toLocaleTimeString();
            let colorClass = '';
            
            if (type === 'ERROR') colorClass = 'text-red-400 font-bold';
            else if (type === 'WARN') colorClass = 'text-yellow-400';
            else if (type === 'SUCCESS') colorClass = 'text-green-400 font-bold';
            else if (type === 'AI_GEMINI') colorClass = 'text-purple-400 font-bold';
            else colorClass = 'text-slate-300';

            const formattedMsg = `[${timestamp}] [${type}] ${msg}`;
            systemLogs.push(formattedMsg);

            // Update UI Console
            const span = document.createElement('span');
            span.className = `block ${colorClass}`;
            span.textContent = formattedMsg;
            logContentEl.appendChild(span);
            logContentEl.scrollTop = logContentEl.scrollHeight;

            console.log(`%c${formattedMsg}`, type === 'ERROR' ? 'color: red' : type === 'WARN' ? 'color: orange' : 'color: inherit');
        }

        async function apiFetch(path, options = {}, label = '') {
            const method = options.method || 'GET';
            addLog(`[API] -> ${method} ${path}${label ? ` (${label})` : ''}`, 'INFO');
            const response = await fetch(`${API_BASE}${path}`, options);
            const raw = await response.text();
            const body = raw || '<empty>';
            const compact = body.replace(/\s+/g, ' ').trim();
            const isVerbosePoll = label === 'answer_poll' || label === 'keypress_poll' || label === 'wait_in_progress';
            const shouldLogResponseBody = !isVerbosePoll || !response.ok;
            const responseMsg = shouldLogResponseBody
                ? `[API] <- ${response.status} ${path} ${compact.slice(0, 300)}`
                : `[API] <- ${response.status} ${path}`;
            addLog(responseMsg, response.ok ? 'SUCCESS' : 'WARN');
            let parsed = null;
            try { parsed = JSON.parse(raw); } catch (e) {}
            return { response, raw, parsed };
        }

        async function loadVersion() {
            let version = 'v0';
            let podName = 'n/a';
            try {
                const { response, parsed } = await apiFetch('/api/version', {}, 'version');
                if (response.ok && parsed?.version) {
                    version = parsed.backend_version || parsed.version;
                    podName = parsed.pod_name || 'n/a';
                }
            } catch (e) {
                addLog(`Version fetch failed: ${e.message}`, 'WARN');
            }

            const versionNode = document.getElementById('appVersion');
            if (versionNode) {
                versionNode.innerText = version;
            }

            const podNode = document.getElementById('podName');
            if (podNode) {
                podNode.innerText = podName;
            }
        }

        function downloadLogs() {
            addLog("Downloading log file...", "INFO");
            const blob = new Blob([systemLogs.join('\n')], { type: 'text/plain' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `voice_agent_logs_${new Date().getTime()}.txt`;
            a.click();
            window.URL.revokeObjectURL(url);
        }

        addLog("Application started.", "SUCCESS");
        loadVersion();

        const formArea = document.getElementById('formArea');
        const callArea = document.getElementById('callArea');
        const contactsFile = document.getElementById('contactsFile');
        const scriptFile = document.getElementById('scriptFile');
        const callDelayInput = document.getElementById('callDelay');
        const voiceSelect = document.getElementById('voiceSelect');
        const twilioAccountSidInput = document.getElementById('twilioAccountSid');
        const twilioAuthTokenInput = document.getElementById('twilioAuthToken');
        const twilioPhoneNumberInput = document.getElementById('twilioPhoneNumber');
        const useServerTwilioCreds = document.getElementById('useServerTwilioCreds');
        const toggleAuthTokenBtn = document.getElementById('toggleAuthTokenBtn');
        const enableRecording = document.getElementById('enableRecording');
        const enableTranscript = document.getElementById('enableTranscript');
        const storageModeSelect = document.getElementById('storageMode');
        const artifactSubfolderInput = document.getElementById('artifactSubfolder');
        const payerProfileSelect = document.getElementById('payerProfileSelect');
        const payerProfileHint = document.getElementById('payerProfileHint');
        const startCallBtn = document.getElementById('startCallBtn');
        const endCallBtn = document.getElementById('endCallBtn');
        const errorBox = document.getElementById('errorBox');
        
        const campaignProgress = document.getElementById('campaignProgress');
        const callStatusText = document.getElementById('callStatusText');
        const targetDisplay = document.getElementById('targetDisplay');
        const contactTypeDisplay = document.getElementById('contactTypeDisplay');
        const propConverterVal = document.getElementById('propConverterVal');
        const propPayerVal = document.getElementById('propPayerVal');
        const callRing = document.getElementById('callRing');
        const avatarIcon = document.getElementById('avatarIcon');
        
        const excelScriptArea = document.getElementById('excelScriptArea');
        const scriptProgress = document.getElementById('scriptProgress');
        const currentQuestionText = document.getElementById('currentQuestion');
        
        const speakQuestionBtn = document.getElementById('speakQuestionBtn');
        const answerYesBtn = document.getElementById('answerYesBtn');
        const answerNoBtn = document.getElementById('answerNoBtn');
        const yesPreview = document.getElementById('yesPreview');
        const noPreview = document.getElementById('noPreview');
        const transcriptBox = document.getElementById('transcriptBox');
        const heardText = document.getElementById('heardText');
        const nextContactBtn = document.getElementById('nextContactBtn');
        const recordingStatus = document.getElementById('recordingStatus');
        const transcriptionStatus = document.getElementById('transcriptionStatus');
        const transcriptSnippet = document.getElementById('transcriptSnippet');
        const artifactStatus = document.getElementById('artifactStatus');
        const insightQuestion = document.getElementById('insightQuestion');
        const insightSource = document.getElementById('insightSource');
        const insightAnswer = document.getElementById('insightAnswer');
        const ivrPhase = document.getElementById('ivrPhase');
        const ivrLastHeard = document.getElementById('ivrLastHeard');
        const ivrLastDigit = document.getElementById('ivrLastDigit');
        const activeMemberIdStatus = document.getElementById('activeMemberIdStatus');
        const remainingMemberIdsStatus = document.getElementById('remainingMemberIdsStatus');

        let currentCallSid = null;
        let callTimeout = null;
        let pollingInterval = null;
        let lastPolledState = null;
        let askAttemptCount = 0;
        let firstQuestionPending = false;
        let customerCareMode = false;
        let lastInsightSignature = '';
        let lastTranscriptCount = 0;
        
        // Campaign Data
        let contactsData = [];
        let scriptData = [];
        let answersDataIndex = {};
        let filteredScriptData = [];
        let currentContactIndex = 0;
        let currentStepIndex = 0;
        let payerProfiles = [];

        loadPayerProfiles();

        async function loadPayerProfiles() {
            try {
                const { response, parsed } = await apiFetch('/api/payer-profiles', {}, 'payer_profiles');
                if (!response.ok) throw new Error(parsed?.error || 'Failed to fetch payer profiles');
                const profiles = Array.isArray(parsed?.profiles) ? parsed.profiles : [];
                payerProfiles = profiles;
                if (!payerProfileSelect) return;
                payerProfileSelect.innerHTML = '';
                if (!profiles.length) {
                    payerProfileSelect.innerHTML = '<option value="">No payer profiles configured</option>';
                    if (payerProfileHint) payerProfileHint.innerText = 'Configure payer_profiles.json with profile name and phone number.';
                    return;
                }
                profiles.forEach((profile) => {
                    const option = document.createElement('option');
                    option.value = String(profile.profile_name || '').trim();
                    const label = String(profile.display_name || option.value).trim();
                    const phone = String(profile.phone_number || '').trim();
                    option.textContent = phone ? `${label} (${phone})` : `${label} (number missing)`;
                    option.disabled = !profile.configured;
                    payerProfileSelect.appendChild(option);
                });
                const firstConfigured = profiles.find((p) => p && p.configured);
                payerProfileSelect.value = String(firstConfigured?.profile_name || profiles[0]?.profile_name || '').trim();
                if (payerProfileHint) payerProfileHint.innerText = 'Campaign will dial selected profile fixed number.';
            } catch (error) {
                if (payerProfileSelect) payerProfileSelect.innerHTML = '<option value="">Failed to load payer profiles</option>';
                if (payerProfileHint) payerProfileHint.innerText = 'Unable to load payer profiles from backend.';
                addLog(`Payer profile load failed: ${error.message}`, 'WARN');
            }
        }

        function updateLiveInsights(state) {
            const data = state || {};
            if (insightQuestion) insightQuestion.innerText = String(data.last_question_from_customer_care || data.last_speech || '-');
            if (insightSource) insightSource.innerText = String(data.last_answer_source || '-');
            if (insightAnswer) insightAnswer.innerText = String(data.last_answer_text || '-');
            if (ivrPhase) ivrPhase.innerText = String(data.ivr_phase || '-').replace(/_/g, ' ');
            if (ivrLastHeard) ivrLastHeard.innerText = String(data.ivr_last_heard_text || '-');
            if (ivrLastDigit) ivrLastDigit.innerText = String(data.ivr_last_pressed_digit || '-');
            const activeMemberId = String(data.active_member_id || data.contact_context?.member_id || '-').trim() || '-';
            if (activeMemberIdStatus) activeMemberIdStatus.innerText = activeMemberId;
            const queue = Array.isArray(data.campaign_contact_queue) ? data.campaign_contact_queue : [];
            const memberIds = queue
                .map((row) => String((row || {}).member_id || '').trim())
                .filter(Boolean);
            if (remainingMemberIdsStatus) {
                remainingMemberIdsStatus.innerText = memberIds.length ? memberIds.join(', ') : '-';
            }
            const totalContacts = Number(data.campaign_total_contacts || contactsData.length || 1);
            const remaining = Number(data.campaign_remaining_contacts_after_current || queue.length || 0);
            const position = Math.max(1, totalContacts - remaining);
            if (campaignProgress) campaignProgress.innerText = `Member ${position} of ${totalContacts}`;
        }

        function emitTranscriptEvents(state) {
            const entries = Array.isArray(state?.interaction_transcript) ? state.interaction_transcript : [];
            if (!entries.length) {
                lastTranscriptCount = 0;
                return;
            }
            if (entries.length < lastTranscriptCount) lastTranscriptCount = 0;
            const recent = entries.slice(lastTranscriptCount);
            recent.forEach((entry) => {
                const source = String(entry?.source || '');
                const text = String(entry?.text || '').trim();
                if (!text) return;
                if (source === 'ivr_menu_prompt') {
                    addLog(`[IVR] Heard: "${text}"`, 'INFO');
                } else if (source === 'ivr_navigation') {
                    addLog(`[IVR] Action: ${text}`, 'SUCCESS');
                } else if (source === 'ivr_navigation_fallback') {
                    addLog(`[IVR] Fallback: ${text}`, 'WARN');
                } else if (source === 'customer_care_intro') {
                    addLog(`[Customer Care] Intro: ${text}`, 'SUCCESS');
                } else if (source === 'customer_care_live_question') {
                    addLog(`[Customer Care] Question detected: "${text}"`, 'INFO');
                }
            });
            lastTranscriptCount = entries.length;
        }

        function normalizeStatusText(value, fallback = 'unknown') {
            const text = String(value || '').trim().toLowerCase();
            if (!text) return fallback;
            if (text === 'disabled') return 'off';
            if (text === 'not_requested') return 'not requested';
            return text;
        }

        function findContactAnswer(contact, step) {
            if (!contact || !step) return null;
            const answerKeys = ['answer', 'Answer', 'response', 'Response', 'info', 'Info'];
            const questionKey = String(step.lookupKey || step.answer_key || step.id || '').trim().toLowerCase();
            if (!questionKey) return null;

            for (const [key, value] of Object.entries(contact)) {
                if (String(key).trim().toLowerCase() === questionKey && String(value || '').trim()) {
                    return String(value).trim();
                }
            }

            for (const k of answerKeys) {
                if (contact[k] && String(contact[k]).trim()) {
                    return String(contact[k]).trim();
                }
            }
            return null;
        }

        function buildMenuPrompt(step) {
            const base = String(step.question || '').trim();
            const options = step.menuOptions || step.menu_options || {};
            const parts = [base];
            if (options && typeof options === 'object') {
                Object.keys(options).sort().forEach((digit) => {
                    const opt = options[digit];
                    if (opt && typeof opt === 'object') {
                        parts.push(`Press ${digit} for ${opt.label || opt.reply || 'option ' + digit}.`);
                    } else {
                        parts.push(`Press ${digit}.`);
                    }
                });
            }
            return parts.join(' ');
        }

        // Parse Contacts
        if (contactsFile) {
            contactsFile.addEventListener('change', (e) => {
                addLog("Contacts file selected, attempting to read...", "INFO");
                const file = e.target.files[0];
                if (!file) { contactsData = []; return; }
                const reader = new FileReader();
                reader.onload = (evt) => {
                    try {
                        const workbook = XLSX.read(evt.target.result, {type: 'binary'});
                        const excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[workbook.SheetNames[0]]);
                        contactsData = excelRows.map(row => ({
                            ...row,
                            phone: String(row['Phone'] || row['phone'] || row['phone_number'] || row['Number'] || '').trim(),
                            type: String(row['Type'] || row['type'] || row['TYPE'] || '').trim(),
                            member_id: String(row['member_id'] || row['Member ID'] || row['MemberID'] || row['memberId'] || row['member id'] || '').trim()
                        })).filter(row => row.member_id !== '');
                        
                        if (contactsData.length > 0) {
                            showError(`Loaded ${contactsData.length} contacts successfully!`, true);
                            addLog(`Successfully parsed ${contactsData.length} contacts from Excel.`, "SUCCESS");
                        } else {
                            showError("No contacts found. Check required column: member_id.");
                        }
                    } catch(err) { 
                        showError("Error reading Contacts file."); 
                    }
                };
                reader.readAsBinaryString(file);
            });
        }

        // Parse Answers Script
        if (scriptFile) {
            scriptFile.addEventListener('change', (e) => {
                addLog("Answers file selected, attempting to parse JSON/YAML...", "INFO");
                const file = e.target.files[0];
                if (!file) { scriptData = []; answersDataIndex = {}; return; }
                
                const reader = new FileReader();
                reader.onload = (evt) => {
                    try {
                        const fileContent = evt.target.result;
                        let parsedData = null;

                        if (file.name.toLowerCase().endsWith('.json')) {
                            parsedData = JSON.parse(fileContent);
                        } else if (file.name.toLowerCase().endsWith('.yaml') || file.name.toLowerCase().endsWith('.yml')) {
                            parsedData = jsyaml.load(fileContent);
                        } else {
                            throw new Error("Unsupported file format. Please upload .json or .yaml");
                        }

                        const normalizeMemberId = (v) => String(v || '').trim();
                        const getMemberId = (row) => {
                            if (!row || typeof row !== 'object') return '';
                            return normalizeMemberId(
                                row.member_id || row['Member ID'] || row.MemberID || row.memberId || row['member id']
                            );
                        };

                        let rows = [];
                        if (Array.isArray(parsedData)) {
                            rows = parsedData.filter(r => r && typeof r === 'object');
                        } else if (parsedData && typeof parsedData === 'object') {
                            if (getMemberId(parsedData)) {
                                rows = [parsedData];
                            } else {
                                rows = Object.values(parsedData).filter(r => r && typeof r === 'object');
                            }
                        }

                        const index = {};
                        rows.forEach((row) => {
                            const memberId = getMemberId(row);
                            if (memberId) {
                                index[memberId] = row;
                            }
                        });

                        answersDataIndex = index;
                        scriptData = rows;

                        const loadedCount = Object.keys(answersDataIndex).length;
                        if (loadedCount > 0) {
                            showError(`Loaded ${loadedCount} member answers successfully!`, true);
                            addLog(`Successfully parsed ${loadedCount} member rows from Answers JSON/YAML.`, "SUCCESS");
                        } else {
                            showError("No valid member answers found. Ensure file has member_id.");
                        }
                    } catch(err) { 
                        showError("Error reading Answers file. Ensure valid JSON/YAML with member_id."); 
                    }
                };
                reader.readAsText(file);
            });
        }

        function showError(msg, isSuccess=false) {
            errorBox.innerText = msg;
            errorBox.className = isSuccess ? "bg-green-50 border-l-4 border-green-500 p-3 rounded-r-md text-green-700 text-sm mb-4 block" : "bg-red-50 border-l-4 border-red-500 p-3 rounded-r-md text-red-700 text-sm mb-4 block";
            errorBox.classList.remove('hidden');
            setTimeout(() => errorBox.classList.add('hidden'), 4000);
        }

        function getCredentialPayload() {
            return {
                account_sid: (twilioAccountSidInput?.value || '').trim(),
                auth_token: (twilioAuthTokenInput?.value || '').trim(),
                from_number: (twilioPhoneNumberInput?.value || '').trim()
            };
        }

        async function fetchServerTwilioConfig() {
            try {
                const { response, parsed } = await apiFetch('/api/debug/twilio-config', {}, 'twilio_config');
                if (!response.ok || !parsed) return null;
                return parsed;
            } catch (e) {
                return null;
            }
        }

        function validateCredentialInputs() {
            if (useServerTwilioCreds?.checked) {
                return null;
            }
            const creds = getCredentialPayload();
            if (!creds.account_sid || !creds.auth_token || !creds.from_number) {
                showError("Please enter Twilio Account SID, Auth Token, and Phone Number.");
                return null;
            }
            if (!/^AC[a-zA-Z0-9]{32}$/.test(creds.account_sid)) {
                showError("Twilio Account SID looks invalid. Expected format: AC followed by 32 chars.");
                return null;
            }
            if (!/^\+[1-9]\d{7,14}$/.test(creds.from_number)) {
                showError("Twilio Phone Number must be in E.164 format, e.g. +12605973908.");
                return null;
            }
            return creds;
        }

        function getEffectiveCredentials() {
            if (useServerTwilioCreds?.checked) {
                return null;
            }
            return validateCredentialInputs();
        }

        function getStorageOptions() {
            const cleanedSubfolder = sanitizeArtifactSubfolder((artifactSubfolderInput?.value || 'default'));
            if (artifactSubfolderInput) artifactSubfolderInput.value = cleanedSubfolder;
            return {
                mode: (storageModeSelect?.value || 'twilio_only'),
                subfolder: cleanedSubfolder
            };
        }

        function getIvrNavigationOptions() {
            return {
                enabled: true,
                department_keywords: [],
                representative_keywords: []
            };
        }

        function sanitizeArtifactSubfolder(value) {
            const trimmed = String(value || '').trim().replace(/\\/g, '/').replace(/^\/+|\/+$/g, '');
            const normalized = trimmed
                .replace(/[^a-zA-Z0-9_./-]/g, '-')
                .replace(/\/+/g, '/');
            const parts = normalized.split('/').filter((part) => part && part !== '.' && part !== '..');
            const safe = parts.join('/');
            return (safe || 'default').slice(0, 120);
        }

        function buildUploadedAnswerBank() {
            const pairs = [];
            const pushPair = (question, answer, memberId = '') => {
                const q = String(question || '').trim();
                const a = String(answer || '').trim();
                if (q && a) pairs.push({ question: q, answer: a, member_id: String(memberId || '').trim() });
            };

            scriptData.forEach((row) => {
                if (!row || typeof row !== 'object') return;
                if (row.question || row.Question) {
                    pushPair(
                        row.question || row.Question,
                        row.answer || row.Answer,
                        row.member_id || row['Member ID'] || row.MemberID || row.memberId || row['member id'] || ''
                    );
                    return;
                }
                const memberId = String(row.member_id || row['Member ID'] || row.MemberID || row.memberId || row['member id'] || '').trim();
                const fields = Object.entries(row)
                    .filter(([k, v]) => String(v || '').trim() && !['phone', 'Phone', 'type', 'Type'].includes(String(k)))
                    .map(([k, v]) => ({ question: `What is ${String(k).replace(/_/g, ' ')}?`, answer: String(v), member_id: memberId }));
                fields.forEach((it) => {
                    const q = String(it.question || '').trim();
                    const a = String(it.answer || '').trim();
                    if (q && a) pairs.push({ question: q, answer: a, member_id: String(it.member_id || '') });
                });
            });

            const seen = new Set();
            return pairs.filter((pair) => {
                const key = `${pair.question.toLowerCase()}::${pair.answer}::${String(pair.member_id || '')}`;
                if (seen.has(key)) return false;
                seen.add(key);
                return true;
            }).slice(0, 500);
        }

        function buildUploadedAnswerBankForMember(memberId) {
            const all = buildUploadedAnswerBank();
            const member = String(memberId || '').trim();
            if (!member) return all;
            return all.filter((item) => {
                const rowMemberId = String(item.member_id || '').trim();
                return rowMemberId === '' || rowMemberId === member;
            });
        }

        function buildRemainingCampaignQueue() {
            const queue = [];
            for (let idx = currentContactIndex + 1; idx < contactsData.length; idx += 1) {
                const row = contactsData[idx] || {};
                const memberId = String(row.member_id || '').trim();
                if (!memberId) continue;
                queue.push({ ...row, member_id: memberId });
            }
            return queue;
        }


        toggleAuthTokenBtn?.addEventListener('click', () => {
            if (!twilioAuthTokenInput) return;
            const showing = twilioAuthTokenInput.type === 'text';
            twilioAuthTokenInput.type = showing ? 'password' : 'text';
            toggleAuthTokenBtn.innerText = showing ? 'View' : 'Hide';
        });


        startCallBtn.addEventListener('click', () => {
            addLog("Start Campaign button clicked.", "INFO");
            
            formArea.classList.add('hidden');
            callArea.classList.remove('hidden'); callArea.classList.add('flex');
            
            currentContactIndex = 0;
            
            if (pollingInterval) clearInterval(pollingInterval);
            currentCallSid = null;
            lastPolledState = null;
            askAttemptCount = 0;
            firstQuestionPending = false;
            customerCareMode = false;
            lastTranscriptCount = 0;
            if (recordingStatus) recordingStatus.innerText = '';
            if (transcriptionStatus) transcriptionStatus.innerText = '';
            if (transcriptSnippet) {
                transcriptSnippet.classList.add('hidden');
                transcriptSnippet.innerText = '';
            }
            if (artifactStatus) artifactStatus.innerText = '';
            updateLiveInsights({});
            lastInsightSignature = '';
            
            campaignProgress.innerText = `Connecting...`;
            targetDisplay.innerText = "Dialing Campaign...";
            contactTypeDisplay.innerText = "";

            filteredScriptData = [];
            currentStepIndex = 0;

            if (nextContactBtn) nextContactBtn.classList.add('hidden');
            if (excelScriptArea) excelScriptArea.classList.remove('hidden');

            callStatusText.innerText = "Calling via Twilio...";
            callRing.className = "absolute inset-0 rounded-full border-4 border-indigo-500 opacity-50 pulse-ring";
            if (currentQuestionText) currentQuestionText.innerText = "Customer care mode: press 9 and ask questions.";
            if (scriptProgress) scriptProgress.innerText = "Live Q&A";
            if (speakQuestionBtn) speakQuestionBtn.classList.add('hidden');
            document.getElementById('answerButtonsRow')?.classList.add('hidden');
            
            makeTwilioCall();
        });

        nextContactBtn?.addEventListener('click', () => {
            addLog("Campaign has ended.", "INFO");
        });

        function triggerNextContactAuto() {
            addLog("Campaign Call Completed.", "SUCCESS");
            if (callStatusText) callStatusText.innerText = "Call completed.";
            if (targetDisplay) targetDisplay.innerText = "Campaign Finished";
            if (contactTypeDisplay) contactTypeDisplay.innerText = "All members have been processed.";
            if (callRing) callRing.className = "absolute inset-0 rounded-full border-4 border-slate-500 opacity-50";
            if (excelScriptArea) excelScriptArea.classList.add('hidden');
            if (nextContactBtn) nextContactBtn.classList.add('hidden');
        }

        async function makeTwilioCall() {
            try {
                const { response, parsed } = await apiFetch('/api/call', {
                    method: 'POST', 
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({})
                }, 'start_call');
                
                const data = parsed || {};
                
                if (response.ok) {
                    currentCallSid = data.call_sid;
                    lastTranscriptCount = 0;
                    if (recordingStatus) {
                        recordingStatus.innerText = data.recording_enabled ? 'enabled' : '';
                    }
                    if (transcriptionStatus) {
                        transcriptionStatus.innerText = data.transcript_enabled ? 'enabled' : '';
                    }
                    if (artifactStatus) {
                        artifactStatus.innerText = `Storage mode: ${data.storage_mode || 'twilio_only'} | Subfolder: ${data.artifact_subfolder || 'default'}`;
                    }
                    if (data.ivr_navigation_enabled) {
                        const categoryNote = data.ivr_matched_category ? ` (${data.ivr_matched_category.replace(/_/g, ' ')})` : '';
                        callStatusText.innerHTML = `Ringing!<br/><span class='text-yellow-400'>Auto-navigating customer care IVR menu${categoryNote}...</span>`;
                        if (data.ivr_matched_category) {
                            addLog(`Auto-detected IVR category: ${data.ivr_matched_category} (source: ${data.ivr_matched_category_source})`, 'INFO');
                        }
                    } else {
                        addLog('Auto IVR navigation is OFF for this call. Waiting for manual press 9.', 'WARN');
                        callStatusText.innerHTML = "Ringing!<br/><span class='text-yellow-400'>Answer & PRESS 9.</span><br/>Waiting for keypress...";
                    }
                    callRing.classList.replace('border-indigo-500', 'border-green-500');
                    startPollingState();
                } else {
                    callStatusText.innerText = "API Error: " + (data.error || "Unknown error");
                    callRing.classList.replace('border-indigo-500', 'border-red-500');
                }
            } catch (error) {
                callStatusText.innerText = "Connection Error. Backend API is not reachable.";
                callRing.classList.replace('border-indigo-500', 'border-red-500');
            }
        }

        function syncTwilioInputState() {
            const disabled = !!useServerTwilioCreds?.checked;
            [twilioAccountSidInput, twilioAuthTokenInput, twilioPhoneNumberInput].forEach((el) => {
                if (!el) return;
                el.disabled = disabled;
                el.classList.toggle('bg-slate-100', disabled);
            });
            if (toggleAuthTokenBtn) {
                toggleAuthTokenBtn.disabled = disabled;
                toggleAuthTokenBtn.classList.toggle('opacity-50', disabled);
            }
            if (disabled) {
                addLog('Using server Twilio credentials from .env for this run.', 'INFO');
                fetchServerTwilioConfig().then((cfg) => {
                    if (!cfg) {
                        addLog('Unable to read server Twilio config. Check backend availability.', 'WARN');
                        return;
                    }
                    addLog(`Server Twilio SID=${cfg.account_sid_masked || ''} phone=${cfg.phone_number || ''} token_set=${cfg.auth_token_set ? 'yes' : 'no'} token_len=${cfg.auth_token_length || 0}`, 'INFO');
                    if (!cfg.auth_token_set || !cfg.phone_number || !cfg.account_sid_masked) {
                        addLog('Server .env Twilio config appears incomplete.', 'ERROR');
                    }
                });
            }
        }

        useServerTwilioCreds?.addEventListener('change', syncTwilioInputState);
        syncTwilioInputState();

        function loadExcelStep() {
            if (currentStepIndex >= filteredScriptData.length) {
                if (currentQuestionText) currentQuestionText.innerText = "✅ Script Complete for this contact!";
                if (speakQuestionBtn) speakQuestionBtn.classList.add('hidden');
                document.getElementById('answerButtonsRow')?.classList.add('hidden');
                if (scriptProgress) scriptProgress.innerText = "Done";
                if (transcriptBox) transcriptBox.classList.add('hidden');
                callRing.className = "absolute inset-0 rounded-full border-4 border-slate-500 opacity-50";
                callStatusText.innerText = "Call flow finished.";
                nextContactBtn.classList.remove('hidden');
                return;
            }
            
            if (speakQuestionBtn) speakQuestionBtn.classList.remove('hidden');
            document.getElementById('answerButtonsRow')?.classList.remove('hidden');
            if (transcriptBox) transcriptBox.classList.add('hidden');
            
            if (answerYesBtn) answerYesBtn.classList.remove('ring-4', 'ring-green-400', 'opacity-50');
            if (answerNoBtn) answerNoBtn.classList.remove('ring-4', 'ring-red-400', 'opacity-50');
            if (answerYesBtn) answerYesBtn.classList.add('opacity-100');
            if (answerNoBtn) answerNoBtn.classList.add('opacity-100');
            
            const step = filteredScriptData[currentStepIndex];
            if (scriptProgress) scriptProgress.innerText = `${currentStepIndex + 1} / ${filteredScriptData.length}`;
            if (currentQuestionText) currentQuestionText.innerText = step.question;
            if (yesPreview) yesPreview.innerText = "Says: " + (step.yes || "...");
            if (noPreview) noPreview.innerText = "Says: " + (step.no || "...");

            if ((step.questionType || '').toLowerCase() === 'menu') {
                if (yesPreview) yesPreview.innerText = 'DTMF Menu';
                if (noPreview) noPreview.innerText = 'Awaiting key press';
            }
        }

        speakQuestionBtn?.addEventListener('click', () => {
            executeAskAndListen();
        });

        async function executeAskAndListen(isAutoRetry = false) {
            if (!currentCallSid) return; 
            if (!filteredScriptData.length || !filteredScriptData[currentStepIndex]) {
                callStatusText.innerText = "Live Q&A mode is active. Press 9 and ask your question.";
                resetSpeakBtn();
                return;
            }

            const step = filteredScriptData[currentStepIndex];
            const voice = voiceSelect.value;
            callStatusText.innerText = "Asking question...";

            const currentContact = contactsData[currentContactIndex] || {};
            const lookupAnswer = findContactAnswer(currentContact, step);
            let effectiveQuestion = step.question;
            let yesResponse = step.yes;
            let noResponse = step.no;

            if ((step.questionType || '').toLowerCase() === 'menu') {
                effectiveQuestion = buildMenuPrompt(step);
                yesResponse = step.yes || 'Thank you. Processing your menu selection.';
                noResponse = step.no || 'No valid selection captured.';
            } else if (lookupAnswer) {
                effectiveQuestion = `${step.question} ${lookupAnswer}`;
                addLog(`Contact sheet answer found for step '${step.id || step.question}'.`, 'SUCCESS');
            } else {
                addLog(`No contact sheet answer found for step '${step.id || step.question}'. AI fallback can handle out-of-bounds questions.`, 'WARN');
            }
            
            if (speakQuestionBtn) {
                speakQuestionBtn.disabled = true;
                speakQuestionBtn.classList.replace('bg-blue-600', 'bg-amber-600');
                speakQuestionBtn.classList.replace('hover:bg-blue-500', 'hover:bg-amber-500');
                speakQuestionBtn.innerHTML = '<i data-lucide="loader" class="w-4 h-4 animate-spin"></i> <span>Listening to customer...</span>';
            }
            lucide.createIcons();

            try {
                const { response, parsed } = await apiFetch('/api/ask_and_listen', {
                    method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ 
                        call_sid: currentCallSid,
                        question: effectiveQuestion,
                        yes_response: yesResponse,
                        no_response: noResponse,
                        voice: voice,
                        menu_options: (step.menuOptions || step.menu_options || {})
                    })
                }, isAutoRetry ? 'ask_retry' : 'ask_primary');
                
                    if (response.ok) {
                        askAttemptCount = 0;
                        firstQuestionPending = false;
                        addLog("Outcome: script_started", "SUCCESS");
                        callStatusText.innerText = "Asking question...";
                        startPollingState();
                    } else {
                    let errorMsg = "Failed to start listening.";
                    let shouldAdvance = false;
                    try {
                        const err = parsed || {};
                        if (err?.code === 'CALL_NOT_ACTIVE') {
                            const state = (err.twilio_status || '').toLowerCase();
                            if (['queued', 'ringing'].includes(state) && askAttemptCount < 3) {
                                askAttemptCount += 1;
                                const delayMs = askAttemptCount * 1000;
                                callStatusText.innerText = `Call not active yet (${state}). Retrying in ${delayMs / 1000}s...`;
                                addLog(`Retrying ask_and_listen attempt ${askAttemptCount} due to status=${state}`, "WARN");
                                setTimeout(() => executeAskAndListen(true), delayMs);
                                return;
                            }
                            errorMsg = `Call ended before script could continue (${state || 'not active'}).`;
                            shouldAdvance = true;
                            addLog(`Outcome: call_not_active_timeout (${state || 'unknown'})`, "WARN");
                        } else if (err?.error) {
                            errorMsg = err.error;
                        }
                    } catch (e) {}
                    callStatusText.innerText = errorMsg;
                    resetSpeakBtn();
                    if (shouldAdvance) {
                        setTimeout(() => {
                            triggerNextContactAuto();
                        }, 1000);
                    }
                }
            } catch (e) {
                callStatusText.innerText = "Error connecting to server.";
                resetSpeakBtn();
            }
        }

        function resetSpeakBtn() {
            if (!speakQuestionBtn) return;
            speakQuestionBtn.disabled = false;
            speakQuestionBtn.classList.replace('bg-amber-600', 'bg-blue-600');
            speakQuestionBtn.classList.replace('hover:bg-amber-500', 'hover:bg-blue-500');
            speakQuestionBtn.innerHTML = '<i data-lucide="play" class="w-4 h-4"></i> <span>Ask & Auto-Listen</span>';
            lucide.createIcons();
        }

        function startKeypressPolling() {
            if (pollingInterval) clearInterval(pollingInterval);
            lastPolledState = null;
            firstQuestionPending = false;
            let pollCount = 0;
            
            pollingInterval = setInterval(async () => {
                pollCount += 1;
                try {
                    const { parsed } = await apiFetch(`/api/call_state/${currentCallSid}`, {}, 'keypress_poll');
                    const data = parsed || {};
                    emitTranscriptEvents(data);

                    if (data.selected_profile_name) {
                        const prof = payerProfiles.find((row) => String(row.profile_name || '').trim() === String(data.selected_profile_name).trim()) || {};
                        if (targetDisplay) {
                            targetDisplay.innerText = prof.phone_number || '-';
                        }
                    }
                    if (data.active_member_id) {
                        let text = `Member ID: ${data.active_member_id}`;
                        if (data.contact_context && data.contact_context['Member Name']) {
                            text += ` (${data.contact_context['Member Name']})`;
                        }
                        if (contactTypeDisplay) {
                            contactTypeDisplay.innerText = text;
                        }
                    }

                    if (propConverterVal) {
                        propConverterVal.innerText = data.run_excel_converter ? "Yes" : "No";
                    }
                    if (propPayerVal) {
                        const profileRaw = data.selected_profile_name || "";
                        propPayerVal.innerText = profileRaw.replace(/_/g, " ");
                    }

                    if (data.status === 'call_ended') {
                        clearInterval(pollingInterval);
                        addLog("Outcome: call_completed_before_script", "WARN");
                        callStatusText.innerText = "Call completed before script started.";
                        setTimeout(() => triggerNextContactAuto(), 1000);
                        return;
                    }

                    if (data.status === 'keypress_detected' && !firstQuestionPending) {
                        firstQuestionPending = true;
                        clearInterval(pollingInterval);
                        callRing.classList.replace('border-yellow-500', 'border-green-500');
                        callStatusText.innerText = "Key 9 Pressed! Waiting for active call state...";
                        waitForInProgressAndStart();
                        return;
                    }

                    if (data.status === 'customer_care_listening') {
                        customerCareMode = true;
                        clearInterval(pollingInterval);
                        callRing.classList.replace('border-yellow-500', 'border-green-500');
                        callStatusText.innerText = 'Customer care mode active. Listening for live questions...';
                        addLog('Switched to customer-care live Q&A mode.', 'SUCCESS');
                        startPollingState();
                        return;
                    }

                    if (data.status === 'ivr_navigating') {
                        clearInterval(pollingInterval);
                        callStatusText.innerText = 'Auto IVR navigation in progress...';
                        addLog('Auto IVR navigation started. Waiting to reach representative menu.', 'INFO');
                        startPollingState();
                        return;
                    }

                    if (data.status === 'waiting_for_keypress' && pollCount === 20) {
                        addLog('No keypress webhook received yet after 20s. Check DTMF input and Twilio webhook reachability.', 'WARN');
                        callStatusText.innerText = 'Still waiting for key press 9... If already pressed, webhook may not be reaching server.';
                    }
                } catch (e) {} 
            }, 1000); 
        }

        function waitForInProgressAndStart() {
            let checks = 0;
            const maxChecks = 12;
            const interval = setInterval(async () => {
                checks += 1;
                try {
                    const { parsed } = await apiFetch(`/api/call_state/${currentCallSid}`, {}, 'wait_in_progress');
                    const data = parsed || {};
                    const twilioState = (data.twilio_call_status || '').toLowerCase();

                    if (data.status === 'call_ended') {
                        clearInterval(interval);
                        addLog("Outcome: call_completed_before_script", "WARN");
                        callStatusText.innerText = "Call completed before script started.";
                        setTimeout(() => triggerNextContactAuto(), 1000);
                        return;
                    }

                    if (twilioState === 'in-progress' || twilioState === 'answered') {
                        clearInterval(interval);
                        askAttemptCount = 0;
                        executeAskAndListen();
                        return;
                    }

                    callStatusText.innerText = `Key 9 detected. Waiting for active call (${twilioState || 'unknown'})...`;
                } catch (e) {}

                if (checks >= maxChecks) {
                    clearInterval(interval);
                    addLog("Outcome: call_not_active_timeout (wait_for_in_progress)", "WARN");
                    callStatusText.innerText = "Timed out waiting for active call after key press.";
                    setTimeout(() => triggerNextContactAuto(), 1000);
                }
            }, 1000);
        }

        function startPollingState() {
            if (pollingInterval) clearInterval(pollingInterval);
            lastPolledState = null;

            pollingInterval = setInterval(async () => {
                try {
                    const { parsed } = await apiFetch(`/api/call_state/${currentCallSid}`, {}, 'answer_poll');
                    const data = parsed || {};
                    emitTranscriptEvents(data);

                    if (data.recording && recordingStatus) {
                        recordingStatus.innerText = normalizeStatusText(data.recording.status, 'unknown');
                    }
                    if (data.transcription && transcriptionStatus) {
                        transcriptionStatus.innerText = normalizeStatusText(data.transcription.status, 'unknown');
                        if (data.transcription.text) {
                            transcriptSnippet.classList.remove('hidden');
                            transcriptSnippet.innerText = `Transcript: ${data.transcription.text}`;
                        }
                    }
                    if (data.artifacts && artifactStatus) {
                        const parts = [];
                        if (data.artifacts.recording_saved) parts.push('recording_saved');
                        if (data.artifacts.twilio_deleted) parts.push('twilio_deleted');
                        if (data.artifacts.transcription_saved_path) parts.push('transcript_saved');
                        if (data.artifacts.interaction_saved_path) parts.push('interaction_saved');
                        if (data.artifacts.audit_results_json_path) parts.push(`audit_json_saved`);
                        if (data.artifacts.audit_json_status) parts.push(data.artifacts.audit_json_status);
                        if (data.artifacts.last_error) parts.push(`error=${data.artifacts.last_error}`);
                        artifactStatus.innerText = parts.join(' | ') || artifactStatus.innerText;
                    }

                    if (data.selected_profile_name) {
                        const prof = payerProfiles.find((row) => String(row.profile_name || '').trim() === String(data.selected_profile_name).trim()) || {};
                        if (targetDisplay) {
                            targetDisplay.innerText = prof.phone_number || '-';
                        }
                    }
                    if (data.active_member_id) {
                        let text = `Member ID: ${data.active_member_id}`;
                        if (data.contact_context && data.contact_context['Member Name']) {
                            text += ` (${data.contact_context['Member Name']})`;
                        }
                        if (contactTypeDisplay) {
                            contactTypeDisplay.innerText = text;
                        }
                    }

                    if (propConverterVal) {
                        propConverterVal.innerText = data.run_excel_converter ? "Yes" : "No";
                    }
                    if (propPayerVal) {
                        const profileRaw = data.selected_profile_name || "";
                        propPayerVal.innerText = profileRaw.replace(/_/g, " ");
                    }

                    if (data.ivr_phase) callStatusText.innerText = `Phase: ${String(data.ivr_phase).replace(/_/g, ' ')}`;

                    const hasNewLastAnswer = !!data.last_answer_text && data.last_answer_text !== (lastPolledState?.last_answer_text || '');
                    updateLiveInsights(data);

                    if (data.status === 'answered' || (data.status === 'customer_care_listening' && hasNewLastAnswer)) {
                        clearInterval(pollingInterval);
                        handleAutoAnswered(data);
                        lastPolledState = data;
                        return;
                    }

                    lastPolledState = data;
                } catch (e) {} 
            }, 1500);
        }

        function handleAutoAnswered(data) {
            callStatusText.innerText = "Response Detected!";
            if (transcriptBox) transcriptBox.classList.remove('hidden');
            if (heardText) heardText.innerText = `"${data.last_speech}"`;

            if (customerCareMode || data.detected_intent === 'customer_care_qna' || data.status === 'customer_care_listening') {
                const source = data.last_answer_source || 'unknown';
                const signature = `${data.last_question_from_customer_care || data.last_speech || ''}::${source}::${data.last_answer_text || ''}::${data.menu_digit || ''}`;
                if (signature !== lastInsightSignature) {
                    addLog(`Customer care question handled using ${source}.`, source === 'uploaded_json' || source === 'contacts_sheet' ? 'SUCCESS' : 'AI_GEMINI');
                    addLog(`Question="${data.last_question_from_customer_care || data.last_speech || ''}"`, 'INFO');
                    addLog(`Answer="${data.last_answer_text || ''}" member_id=${data.active_member_id || data.contact_context?.member_id || ''} digit=${data.menu_digit || ''}`, 'INFO');
                    if (data.answer_consistency && data.answer_consistency.status === 'warning') {
                        addLog(`Answer consistency warning: ${(data.answer_consistency.flags || []).join(', ') || 'unknown'}`, 'WARN');
                    }
                    lastInsightSignature = signature;
                }
                callStatusText.innerText = `Answered (${source}). Waiting...`;
                setTimeout(() => startPollingState(), 1200);
                return;
            }
            
            if (data.detected_intent === 'yes' || data.detected_intent === 'no') {
                let resolvedAction = 'next'; 
                
                if (data.detected_intent === 'yes') {
                    if (answerYesBtn) answerYesBtn.classList.add('ring-4', 'ring-green-400');
                    if (answerNoBtn) answerNoBtn.classList.add('opacity-50');
                    resolvedAction = filteredScriptData[currentStepIndex].actionYes;
                } else {
                    if (answerNoBtn) answerNoBtn.classList.add('ring-4', 'ring-red-400');
                    if (answerYesBtn) answerYesBtn.classList.add('opacity-50');
                    resolvedAction = filteredScriptData[currentStepIndex].actionNo;
                }

                resetSpeakBtn();
                addLog(`In-bounds logic. Action defined as: '${resolvedAction}'`, "INFO");
                
                setTimeout(() => {
                    processStepAction(resolvedAction, true);
                }, 4500);
            } else {
                if (data.detected_intent === 'menu') {
                    addLog(`DTMF menu selection detected: digit=${data.menu_digit || ''} next_action=${data.next_action || ''}`, 'INFO');
                    const action = (data.next_action || '').trim().toLowerCase();
                    const fallback = filteredScriptData[currentStepIndex].actionYes || 'next';
                    resetSpeakBtn();
                    setTimeout(() => {
                        processStepAction(action || fallback, true);
                    }, 2500);
                    return;
                }
                // Out of Bounds -> Trigger Dynamic Gemini Speech
                if (heardText) heardText.innerText = `"${data.last_speech}" (Question Detected!)`;
                if (answerYesBtn) answerYesBtn.classList.add('opacity-50');
                if (answerNoBtn) answerNoBtn.classList.add('opacity-50');
                
                handleDynamicQuestion(data.last_speech, filteredScriptData[currentStepIndex]);
            }
        }

        // ==========================================
        // NEW: GEMINI DYNAMIC SPEECH & SEARCH INTEGRATION
        // ==========================================
        async function fetchGeminiWithBackoff(promptText) {
            const apiKey = "AIzaSyD8ysvkrDiNAGFuHa7BK-0ZBnfUzUa0M3w"; // API Key handled by runtime execution environment
            const payload = { 
                contents: [{ parts: [{ text: promptText }] }], 
                systemInstruction: { parts: [{ text: "You are a friendly customer care voice agent. Keep answers conversational, very brief (1-2 sentences), and suitable for text-to-speech reading. Do not use markdown." }] },
                tools: [{ "google_search": {} }]
            };
            
            let delays = [1000, 2000, 4000, 8000, 16000];
            let attempt = 0;
            
            while (attempt < 6) {
                try {
                    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=${apiKey}`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(payload)
                    });
                    
                    if (!response.ok) throw new Error("API responded with status: " + response.status);
                    const result = await response.json();
                    return result.candidates?.[0]?.content?.parts?.[0]?.text;
                } catch (e) {
                    if (attempt < 5) {
                        await new Promise(r => setTimeout(r, delays[attempt]));
                        attempt++;
                    } else {
                        throw new Error("Failed after 5 retries.");
                    }
                }
            }
        }

        async function handleDynamicQuestion(userSpeech, currentStep) {
            callStatusText.innerHTML = "Generating AI Response via <span class='text-purple-400 font-bold'>Gemini Search...</span>";
            avatarIcon.classList.add('ai-pulse');
            avatarIcon.innerHTML = '<i data-lucide="sparkles" class="w-8 h-8 text-purple-400"></i>';
            lucide.createIcons();
            
            addLog(`Out of bounds speech detected. Calling Gemini AI Search...`, "AI_GEMINI");
            
            const promptText = `You are a helpful and polite customer care agent on a phone call. 
The customer was just asked this yes/no script question: "${currentStep.question}"
Instead of answering yes or no, the customer said/asked: "${userSpeech}"

Instructions:
1. Answer their question directly and briefly (1 to 2 short sentences).
2. If you need factual data or up-to-date information to answer their question, use the google_search tool.
3. End your response by politely asking the original question again so they can answer it: "${currentStep.question}"`;

            try {
                const aiText = await fetchGeminiWithBackoff(promptText);
                
                if (!aiText) throw new Error("Empty response from Gemini.");
                
                addLog(`Gemini Output: ${aiText}`, "AI_GEMINI");
                callStatusText.innerText = "Agent Speaking AI Response...";
                
                // Push the new custom question to Twilio to overwrite the "Please hold" fallback
                const { response: twilioRes, parsed: twilioErr } = await apiFetch('/api/ask_and_listen', {
                    method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ 
                        call_sid: currentCallSid, 
                        question: aiText,
                        yes_response: currentStep.yes, 
                        no_response: currentStep.no,
                        voice: voiceSelect.value
                    })
                }, 'ask_dynamic_ai');
                
                avatarIcon.classList.remove('ai-pulse');
                avatarIcon.innerHTML = '<i data-lucide="user" class="w-8 h-8 text-slate-400"></i>';
                lucide.createIcons();

                if (twilioRes.ok) {
                    addLog("Successfully resumed listening with AI-generated question.", "SUCCESS");
                    startPollingState();
                } else {
                    if (twilioErr?.code === 'CALL_NOT_ACTIVE') {
                        addLog(`Outcome: call_not_active_timeout (${twilioErr.twilio_status || 'unknown'})`, "WARN");
                        setTimeout(() => triggerNextContactAuto(), 1000);
                    }
                    addLog("Failed to push AI response to Twilio.", "ERROR");
                    resetSpeakBtn();
                }
                
            } catch (error) {
                addLog(`Gemini API Error: ${error.message}. Repeating original question.`, "ERROR");
                callStatusText.innerText = "AI Error. Repeating question...";
                
                avatarIcon.classList.remove('ai-pulse');
                avatarIcon.innerHTML = '<i data-lucide="user" class="w-8 h-8 text-slate-400"></i>';
                lucide.createIcons();

                // Fallback: Just repeat the question via our normal execute Ask and Listen
                setTimeout(() => {
                    executeAskAndListen(false);
                }, 2000);
            }
        }
        // ==========================================

        function processStepAction(action, autoTrigger) {
            if (action === 'end') {
                triggerNextContactAuto();
            } else if (action === 'next') {
                currentStepIndex++;
                finalizeStepMove(autoTrigger);
            } else {
                const targetIndex = filteredScriptData.findIndex(s => s.id === action);
                if (targetIndex !== -1) {
                    currentStepIndex = targetIndex;
                    finalizeStepMove(autoTrigger);
                } else {
                    currentStepIndex++;
                    finalizeStepMove(autoTrigger);
                }
            }
        }

        function finalizeStepMove(autoTrigger) {
            loadExcelStep();
            if (currentStepIndex < filteredScriptData.length && currentCallSid && autoTrigger) {
                setTimeout(() => executeAskAndListen(), 500);
            }
        }

        // --- MANUAL OVERRIDES ---
        answerYesBtn?.addEventListener('click', async () => {
            if (pollingInterval) clearInterval(pollingInterval);
            resetSpeakBtn();
            
            const step = filteredScriptData[currentStepIndex];
            if (step.yes) await sendManualSpeech(step.yes);
            
            setTimeout(() => {
                processStepAction(step.actionYes, false);
            }, 4500);
        });

        answerNoBtn?.addEventListener('click', async () => {
            if (pollingInterval) clearInterval(pollingInterval);
            resetSpeakBtn();
            
            const step = filteredScriptData[currentStepIndex];
            if (step.no) await sendManualSpeech(step.no);
            
            setTimeout(() => {
                processStepAction(step.actionNo, false);
            }, 4500);
        });

        endCallBtn.addEventListener('click', async () => {
            if (pollingInterval) clearInterval(pollingInterval);
            clearTimeout(callTimeout);
            
            const callSidToEnd = currentCallSid;
            currentCallSid = null;
            
            callStatusText.innerText = "Campaign Aborted...";
            callRing.className = "absolute inset-0 rounded-full border-4 border-slate-500 opacity-50";
            
            if (callSidToEnd) {
                try {
                    await fetch(`/api/call/${callSidToEnd}/end`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' }
                    });
                } catch (err) {
                    console.error('Failed to end call via API:', err);
                }
            }
            
            setTimeout(() => {
                callArea.classList.add('hidden'); callArea.classList.remove('flex');
                formArea.classList.remove('hidden');
                nextContactBtn.classList.add('hidden');
            }, 1000);
        });

document.getElementById('logOpenBtn')?.addEventListener('click', () => {
    document.getElementById('logModal').classList.remove('hidden');
});

document.getElementById('logCloseBtn')?.addEventListener('click', () => {
    document.getElementById('logModal').classList.add('hidden');
});

        document.getElementById('downloadLogsBtn')?.addEventListener('click', () => {
            downloadLogs();
        });
