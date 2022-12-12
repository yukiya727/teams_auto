import pip

if __name__ == '__main__':
    with open('requirements.txt') as f:
        for line in f:
            package = line.strip()
            pip.main(['install', package])
