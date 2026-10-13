from read_check import SafeError
def read_call(client, method, params):
    return client.call(method, params)
