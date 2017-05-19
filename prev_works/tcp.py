#!C:\Python34

# import socket

# class MySocket:
#     """demonstration class only
#       - coded for clarity, not efficiency
#     """

#     def __init__(self, sock=None):
#         if sock is None:
#             self.sock = socket.socket(
#                             socket.AF_INET, socket.SOCK_STREAM)
#         else:
#             self.sock = sock

#     def connect(self, host, port):
#         self.sock.connect((host, port))

#     def mysend(self, msg):
#         totalsent = 0
#         self.sock.send(msg)
#         # while totalsent < len (msg):
#         #     sent = self.sock.send(msg[totalsent:])
#         #     if sent == 0:
#         #         raise RuntimeError("socket connection broken")
#         #     totalsent = totalsent + sent

#     def myreceive(self):
#         chunks = []
#         bytes_recd = 0
#         while bytes_recd < MSGLEN:
#             chunk = self.sock.recv(min(MSGLEN - bytes_recd, 2048))
#             if chunk == b'':
#                 raise RuntimeError("socket connection broken")
#             chunks.append(chunk)
#             bytes_recd = bytes_recd + len(chunk)
#         return b''.join(chunks)

#     def close(self):
#     	self.sock.shutdown()
#     	self.sock.close()


# soc = MySocket()

# soc.connect("localhost", 30000)
# soc.mysend("test")
# soc.close

import socket               # Import socket module
import time

s = socket.socket()         # Create a socket object
host = socket.gethostname() # Get local machine name
port = 30000                # Reserve a port for your service.

s.connect((host, port))
for i in range (1,2):
	s.send(b"test")
	time.sleep(10)
print ("Send...")
# s.shutdown(1)
s.close() # Close the socket when done

