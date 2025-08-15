from flask import Flask, render_template, request, redirect, url_for, session, send_file, jsonify
import pandas as pd
import speech_recognition as sr
import pyttsx3
import os
from fuzzywuzzy import process

app = Flask(__name__)
app.secret_key = 'your_secret_key'

# Load component data from Excel
def load_component_data(filename):
    try:
        data = pd.read_excel(filename)
        component_dict = {}
        for _, row in data.iterrows():
            name = str(row['Name']).strip().lower()
            component_dict[name] = {
                "image": str(row['Image']).strip() if pd.notna(row['Image']) else 'placeholder.jpg',
                "description": str(row['Description']).strip() if pd.notna(row['Description']) else 'No description available.',
                "units": str(row['Units']).strip() if pd.notna(row['Units']) else 'N/A',
                "advantage": str(row['Advantage']).strip() if pd.notna(row['Advantage']) else 'N/A',
                "disadvantage": str(row['Disadvantage']).strip() if pd.notna(row['Disadvantage']) else 'N/A',
                "applications": str(row['Applications']).strip() if pd.notna(row['Applications']) else 'N/A',
                "materials": str(row['Materials']).strip() if pd.notna(row['Materials']) else 'N/A',
                "power": row['Power'] if 'Power' in row and pd.notna(row['Power']) else 'Not available',
                "voltage": row['Voltage'] if 'Voltage' in row and pd.notna(row['Voltage']) else 'Not available',
                "current": row['Current'] if 'Current' in row and pd.notna(row['Current']) else 'Not available',
                "category": row['Category'] if 'Category' in row and pd.notna(row['Category']) else 'Not available'
            }
        print("DEBUG: Loaded component keys:", list(component_dict.keys()))
        return component_dict
    except Exception as e:
        print(f"Error loading dataset: {e}")
        return {}

# Dataset path
current_directory = os.path.dirname(__file__)
dataset_path = os.path.join(current_directory, 'cleaned_component_data.xlsx')
component_data = load_component_data(dataset_path)

# Text-to-Speech helper
def speak(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username and password:
            session['username'] = username
            return redirect(url_for('component'))
    return render_template('signin.html')

@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == 'POST':
        return redirect(url_for('login'))
    return render_template('signup.html')

@app.route('/signin', methods=['GET', 'POST'])
def signin():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username and password:
            session['username'] = username
            return redirect(url_for('component'))
    return render_template('signin.html')

@app.route('/component')
def component():
    if 'username' not in session:
        return redirect(url_for('signin'))
    return render_template('component.html', email=session.get('username', 'User'))

@app.route('/search', methods=['POST'])
def search():
    """Search for components by text with fuzzy matching."""
    name = request.form.get('typed_component', '').strip().lower()
    print(f"DEBUG: User searched for: '{name}'")

    # Exact match
    if name in component_data:
        component = component_data[name]
        return render_template('component_details.html', component=component)

    # Fuzzy match
    best_match, score = process.extractOne(name, component_data.keys())
    print(f"DEBUG: Fuzzy match result: '{best_match}' with score: {score}")

    if score >= 85:
        component = component_data[best_match]
        return render_template('component_details.html', component=component)

    return render_template('component.html', email=session.get('username', 'User'), error="Component not found.")

@app.route('/component-details')
def component_details():
    name = request.args.get('name', '').strip().lower()
    if name in component_data:
        component = component_data[name]
        return render_template('component_details.html', component=component)
    return "Component not found."

@app.route('/recognize-speech', methods=['GET'])
def recognize_speech():
    recognizer = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            print("DEBUG: Adjusting for ambient noise...")
            recognizer.adjust_for_ambient_noise(source, duration=5)
            print("DEBUG: Listening for speech...")
            audio = recognizer.listen(source, timeout=10)
            print("DEBUG: Processing speech...")
            spoken_text = recognizer.recognize_google(audio).lower().strip()
            print(f"DEBUG: Recognized Speech: '{spoken_text}'")

            matched_name = next((name for name in component_data if name == spoken_text or spoken_text in name), None)

            if not matched_name:
                print("DEBUG: Attempting fuzzy matching...")
                matched_name, score = process.extractOne(spoken_text, component_data.keys())
                print(f"DEBUG: Fuzzy matched component: '{matched_name}' with score: {score}")
                if score < 85:
                    matched_name = None

            if matched_name:
                component = component_data[matched_name]
                return jsonify({
                    "status": "success",
                    "component": matched_name,
                    "details": component
                })
            else:
                return jsonify({
                    "status": "error",
                    "message": "No matching component found in the dataset."
                })

    except sr.UnknownValueError:
        return jsonify({"status": "error", "message": "Could not understand your speech. Please try again."})
    except sr.RequestError as e:
        return jsonify({"status": "error", "message": f"Speech recognition service error: {e}"})
    except Exception as e:
        return jsonify({"status": "error", "message": f"An unexpected error occurred: {e}"})

@app.route('/reload-dataset', methods=['POST', 'GET'])
def reload_dataset():
    global component_data
    component_data = load_component_data(dataset_path)
    print("DEBUG: Dataset reloaded.")
    return jsonify({"status": "Dataset reloaded successfully!"})

@app.route('/generate_dataset')
def generate_dataset():
    filename = "electronic_components.csv"
    return send_file(filename, as_attachment=True)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('home'))

if __name__ == '__main__':
    app.run(debug=True)