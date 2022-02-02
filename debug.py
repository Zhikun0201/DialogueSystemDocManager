def not_valid(data):
    if data:
        return False
    data = locals()
    print(f"{data} is not valide")
    return True
