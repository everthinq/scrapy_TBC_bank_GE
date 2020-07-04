from scrapy.cmdline import execute

def main():
    # execute(['scrapy', 'crawl', 'TBC_bank', '--nolog'])
    execute(['scrapy', 'crawl', 'TBC_bank'])

if __name__ == '__main__':
    main()