from flask import Flask, render_template, request, redirect, url_for, session, send_file, jsonify
import pandas as pd
import speech_recognition as sr
import pyttsx3
import os

app = Flask(__name__)
app.secret_key = 'your_secret_key'

# Load component data from the uploaded Excel file
def load_component_data(filename):
    """Load components data from the uploaded Excel file with debugging."""
    try:
        data = pd.read_excel(filename)
        component_dict = {}
        for _, row in data.iterrows():
            component_dict[row['Name'].strip().lower()] = {
                "image": row['Image'].strip() if pd.notna(row['Image']) else 'placeholder.jpg',
                "description": row['Description'].strip() if pd.notna(row['Description']) else 'No description available.',
                "units": row['Units'].strip() if pd.notna(row['Units']) else 'N/A',
                "advantage": row['Advantage'].strip() if pd.notna(row['Advantage']) else 'N/A',
                "disadvantage": row['Disadvantage'].strip() if pd.notna(row['Disadvantage']) else 'N/A',
                "applications": row['Applications'].strip() if pd.notna(row['Applications']) else 'N/A',
                "materials": row['Materials'].strip() if pd.notna(row['Materials']) else 'N/A',
                "power": row['Power'] if 'Power' in row and pd.notna(row['Power']) else 'Not available',
                "voltage": row['Voltage'] if 'Voltage' in row and pd.notna(row['Voltage']) else 'Not available',
                "current": row['Current'] if 'Current' in row and pd.notna(row['Current']) else 'Not available',
                "category": row['Category'] if 'Category' in row and pd.notna(row['Category']) else 'Not available'
            }
        # Debug: Log the loaded dataset keys
        print("DEBUG: Loaded component keys:")
        for key in component_dict.keys():
            print(f"Key: '{key}'")
        return component_dict
    except Exception as e:
        print(f"Error loading dataset: {e}")
        return {}

# Set the path to the component dataset
current_directory = os.path.dirname(__file__)
dataset_path = os.path.join(current_directory, 'cleaned_component_data.xlsx')

# Load components from the uploaded Excel file
component_data = load_component_data(dataset_path)

# Text-to-Speech helper
def speak(text):
    """Text-to-speech feedback."""
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

@app.route('/')
def home():
    """Render the home page."""
    return render_template('home.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Handle user login."""
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username and password:
            session['username'] = username
            return redirect(url_for('component'))
    return render_template('signin.html')

@app.route('/signup', methods=['GET', 'POST'])
def signup():
    """Handle user signup."""
    if request.method == 'POST':
        return redirect(url_for('login'))
    return render_template('signup.html')

@app.route('/signin', methods=['GET', 'POST'])
def signin():
    """Handle user signin."""
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username and password:
            session['username'] = username
            return redirect(url_for('component'))
    return render_template('signin.html')

@app.route('/component')
def component():
    """Render the component search page."""
    if 'username' not in session:
        return redirect(url_for('signin'))
    return render_template('component.html', email=session.get('username', 'User'))

@app.route('/logout')
def logout():
    """Handle user logout."""
    session.clear()
    return redirect(url_for('home'))

@app.route('/recognize-speech', methods=['GET'])
def recognize_speech():
    """Recognize speech with enhanced debugging for dataset and matching."""
    recognizer = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            print("DEBUG: Adjusting for ambient noise...")
            recognizer.adjust_for_ambient_noise(source, duration=5)  # Extended adjustment time
            
            print("DEBUG: Listening for speech...")
            audio = recognizer.listen(source, timeout=10)  # Increased timeout
            
            print("DEBUG: Processing speech...")
            spoken_text = recognizer.recognize_google(audio).lower().strip()
            print(f"DEBUG: Recognized Speech: '{spoken_text}'")
            
            # Debug: Validate recognized speech
            if not spoken_text:
                print("DEBUG: No speech recognized. Empty input detected.")
                return jsonify({
                    "status": "error",
                    "message": "No speech detected. Please try again."
                })
            
            # Debug: Matching Logic
            print("DEBUG: Attempting exact and partial matches...")
            matched_name = next(
                (name for name in component_data.keys() if name == spoken_text or spoken_text in name), None
            )
            
            # If no match found, use fuzzy matching
            if not matched_name:
                print("DEBUG: Attempting fuzzy matching...")
                from fuzzywuzzy import process
                matched_name, score = process.extractOne(spoken_text, component_data.keys())
                print(f"DEBUG: Fuzzy matched component: '{matched_name}' with score: {score}")
                if score < 85:  # Similarity threshold
                    matched_name = None
            
            if matched_name:
                print(f"Matched Component: {matched_name}")
                component = component_data[matched_name]
                return jsonify({
                    "status": "success",
                    "component": matched_name,
                    "details": component
                })
            else:
                print("DEBUG: No matching component found.")
                return jsonify({
                    "status": "error",
                    "message": "No matching component found in the dataset."
                })
    except sr.UnknownValueError:
        print("DEBUG: Could not understand the audio.")
        return jsonify({
            "status": "error",
            "message": "Could not understand your speech. Please try again."
        })
    except sr.RequestError as e:
        print(f"DEBUG: Speech recognition service error: {e}")
        return jsonify({
            "status": "error",
            "message": f"Speech recognition service error: {e}"
        })
    except Exception as e:
        print(f"DEBUG: Unexpected error: {e}")
        return jsonify({
            "status": "error",
            "message": f"An unexpected error occurred: {e}"
        })

@app.route('/search', methods=['POST'])
def search():
    """Search for components by text and display dynamically."""
    name = request.form.get('typed_component', '').lower()
    if name in component_data:
        component = component_data[name]
        return render_template('component_details.html', component=component)
    return "Component not found."

@app.route('/component-details')
def component_details():
    """Render detailed information about a component."""
    name = request.args.get('name', '').lower()
    if name in component_data:
        component = component_data[name]
        return render_template('component_details.html', component=component)
    return "Component not found."

@app.route('/reload-dataset', methods=['POST', 'GET'])
def reload_dataset():
    """Reload the dataset dynamically."""
    global component_data
    component_data = load_component_data(dataset_path)
    print("DEBUG: Dataset reloaded. Available components:", list(component_data.keys()))
    return jsonify({"status": "Dataset reloaded successfully!"})

@app.route('/generate_dataset')
def generate_dataset():
    """Generate and download the dataset."""
    filename = "electronic_components.csv"
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
