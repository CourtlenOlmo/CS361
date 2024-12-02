import time
import zmq
import random
def server_main():
    """Main function to handle ZeroMQ server and book search microservice."""
    context = zmq.Context()
    socket = context.socket(zmq.REP)
    socket.bind("tcp://0.0.0.0:5558")

    print(">>> Server running, waiting for client requests...")
    file = open("quotes.txt", encoding="utf8")
    content = file.readlines()

    while True:
        time.sleep(3)
        line = random.randrange(0,50)
        request = socket.recv()  # Receive goal as string
        prompt = request.decode()
        if prompt == "printBook":
            quote = content[line]
            socket.send_string(quote)

        file.close

        # Destroy and exit
    context.destroy()


if __name__ == "__main__":
    server_main()