from flask import Flask
 
 
import webview
import sys
import threading
 
app = Flask(__name__)

# Define a route for the default URL, which loads the form
@app.route('/')
def form():
    return render_template('form_submit.html')

# Define a route for the action of the form, for example '/hello/'
# We are also defining which type of requests this route is 
# accepting: POST requests in this case
@app.route('/hello/', methods=['POST'])
def hello():
    name=request.form['yourname']
    email=request.form['youremail']
    return render_template('form_action.html', name=name, email=email)
 
def start_server():
    app.run( 
        host="0.0.0.0",
        port=int("80")
  )



if __name__ == '__main__':
    """  https://github.com/r0x0r/pywebview/blob/master/examples/http_server.py
    """
 
    t = threading.Thread(target=start_server)
    t.daemon = True
    t.start()
 
    webview.create_window("It works, Joe!", "http://127.0.0.1:5000/")
 
    sys.exit()