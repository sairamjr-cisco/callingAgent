from flask import Flask, request, jsonify, Response
from flask_cors import CORS
from twilio.rest import Client
import os
import html

app = Flask(__name__)
CORS(app) 

# Replace these with your actual Twilio credentials
TWILIO_ACCOUNT_SID = os.environ.get('TWILIO_ACCOUNT_SID', '')
TWILIO_AUTH_TOKEN = os.environ.get('TWILIO_AUTH_TOKEN', '')
TWILIO_PHONE_NUMBER = os.environ.get('TWILIO_PHONE_NUMBER', '')

client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)

# In-memory dictionary to store the state of active calls
active_calls_state = {}

@app.route('/api/call', methods=['POST'])
def make_call():
    data = request.json
    to_number = data.get('number')
    public_url = data.get('public_url', '').strip()

    if not to_number:
        return jsonify({'error': 'Phone number is required.'}), 400

    public_url = public_url.rstrip('/')
    if public_url and not public_url.startswith('http'):
        public_url = 'https://' + public_url

    try:
        # If we have a public URL, setup a Gather to detect key press 9 before proceeding
        if public_url:
            twiml_instructions = f'''
            <Response>
                <Gather numDigits="1" action="{public_url}/api/webhook/keypress" method="POST" timeout="60">
                    <Say voice="alice">Please press 9 to begin the automated script.</Say>
                </Gather>
                <Pause length="3600"/>
            </Response>
            '''
        else:
            twiml_instructions = '''
            <Response>
                <Say voice="alice">Agent connected. Waiting for operator.</Say>
                <Pause length="3600"/>
            </Response>
            '''

        call = client.calls.create(
            twiml=twiml_instructions,
            to=to_number,
            from_=TWILIO_PHONE_NUMBER
        )
        
        # Initialize state for this call
        active_calls_state[call.sid] = {'status': 'waiting_for_keypress', 'last_speech': ''}

        return jsonify({'message': 'Call initiated successfully', 'call_sid': call.sid}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/webhook/keypress', methods=['POST'])
def handle_keypress():
    """Twilio hits this endpoint when the user presses a key on the initial connection."""
    call_sid = request.form.get('CallSid')
    digits = request.form.get('Digits')
    
    if digits == '9':
        # Update the state so the frontend knows to start injecting the Excel script
        if call_sid in active_calls_state:
            active_calls_state[call_sid]['status'] = 'keypress_detected'
        
        # Pause to keep the line open while the frontend sends the first question
        twiml_response = '''
        <Response>
            <Pause length="3600"/>
        </Response>
        '''
        return Response(twiml_response, mimetype='text/xml')
    else:
        # If they pressed the wrong key, prompt them again
        twiml_response = '''
        <Response>
            <Gather numDigits="1" action="/api/webhook/keypress" method="POST" timeout="60">
                <Say voice="alice">Invalid key. Please press 9 to begin.</Say>
            </Gather>
        </Response>
        '''
        return Response(twiml_response, mimetype='text/xml')


@app.route('/api/ask_and_listen', methods=['POST'])
def ask_and_listen():
    data = request.json
    call_sid = data.get('call_sid')
    question = data.get('question', '')
    yes_response = data.get('yes_response', '')
    no_response = data.get('no_response', '')
    public_url = data.get('public_url', '').strip()

    if not public_url:
        return jsonify({'error': 'Public URL is required for speech recognition.'}), 400

    # SAFETY FIX: Clean trailing slashes to prevent Flask 308 redirects that drop Twilio webhooks
    public_url = public_url.rstrip('/')
    if not public_url.startswith('http'):
        public_url = 'https://' + public_url

    # SAFETY FIX: Escape text to prevent special characters (like '&' or '<') from breaking Twilio's XML
    safe_question = html.escape(question)

    # Store the expected answers in our state
    active_calls_state[call_sid] = {
        'status': 'listening',
        'yes_response': yes_response,
        'no_response': no_response,
        'last_speech': '',
        'detected_intent': ''
    }

    try:
        # Tell Twilio to speak the question, then listen for a reply
        # and POST the transcript to our public URL webhook
        twiml_instructions = f'''
        <Response>
            <Gather input="speech" action="{public_url}/api/webhook/speech" method="POST" timeout="5" speechTimeout="auto">
                <Say voice="alice">{safe_question}</Say>
            </Gather>
            <Say voice="alice">I didn't catch that. Please hold.</Say>
            <Pause length="3600"/>
        </Response>
        '''
        client.calls(call_sid).update(twiml=twiml_instructions)
        return jsonify({'message': 'Listening for response...'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/webhook/speech', methods=['POST'])
def handle_speech():
    """Twilio hits this endpoint when it hears the user speak."""
    call_sid = request.form.get('CallSid')
    speech_result = request.form.get('SpeechResult', '').lower()
    
    state = active_calls_state.get(call_sid, {})
    yes_text = state.get('yes_response', 'Okay.')
    no_text = state.get('no_response', 'Alright.')

    # Logic to match keywords in what the user said
    yes_keywords = ['yes', 'yeah', 'sure', 'yep', 'ok', 'okay', 'correct', 'right', 'please', 'course', 'absolutely']
    no_keywords = ['no', 'nope', 'not', 'wrong', 'incorrect', 'stop']

    intent = 'unknown'
    reply_text = "I'm sorry, I didn't understand if that was a yes or a no. Please hold."

    if any(word in speech_result for word in yes_keywords):
        intent = 'yes'
        reply_text = yes_text
    elif any(word in speech_result for word in no_keywords):
        intent = 'no'
        reply_text = no_text

    # Update our server state so the frontend can display it
    if call_sid in active_calls_state:
        active_calls_state[call_sid]['status'] = 'answered'
        active_calls_state[call_sid]['last_speech'] = speech_result
        active_calls_state[call_sid]['detected_intent'] = intent

    # SAFETY FIX: Escape the reply text as well
    safe_reply = html.escape(reply_text)

    # Tell Twilio what to say back to the user automatically
    twiml_response = f'''
    <Response>
        <Say voice="alice">{safe_reply}</Say>
        <Pause length="3600"/>
    </Response>
    '''
    return Response(twiml_response, mimetype='text/xml')


@app.route('/api/call_state/<call_sid>', methods=['GET'])
def get_call_state(call_sid):
    """Frontend polls this to see if the user answered."""
    state = active_calls_state.get(call_sid, {})
    return jsonify(state), 200

# Keep standard text-to-speech for manual overriding
@app.route('/api/speak', methods=['POST'])
def speak_in_call():
    data = request.json
    call_sid = data.get('call_sid')
    text = data.get('text', '')
    safe_text = html.escape(text)
    try:
        twiml = f'<Response><Say voice="alice">{safe_text}</Say><Pause length="3600"/></Response>'
        client.calls(call_sid).update(twiml=twiml)
        return jsonify({'message': 'Text spoken successfully'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    print("Starting Twilio Call Server on http://127.0.0.1:5001")
    app.run(debug=True, port=5001)