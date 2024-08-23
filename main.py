from src.generator import Generator


if __name__ == '__main__':
    generator = Generator()

    print(f'Json output file: {generator.run()}')
