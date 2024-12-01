import time
import zmq

def server_main():
    """Main function to handle ZeroMQ server and book search microservice."""
    context = zmq.Context()
    socket = context.socket(zmq.REP)
    socket.bind("tcp://0.0.0.0:5556")

    print(">>> Server running, waiting for client requests...")
    file = open("goalTrackText.txt", "a")

    while True:
        time.sleep(3)
        request = socket.recv()  # Receive goal as string
        goal = request.decode()
        print(f">>> Client requested: {goal} hours")

        file.write(goal)
        file.write("\n")
        file.close



        # Destroy and exit
    context.destroy()


if __name__ == "__main__":
    server_main()
