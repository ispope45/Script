import sys
import os


def printProgress(iteration, total, prefix='', suffix='', decimals=1, barLength=100):
    formatStr = "{0:." + str(decimals) + "f}"
    percent = formatStr.format(100 * (iteration / float(total)))
    filledLength = int(round(barLength * iteration / float(total)))
    bar = '#' * filledLength + '-' * (barLength - filledLength)
    sys.stdout.write('\r%s |%s| %s%s %s' % (prefix, bar, percent, '%', suffix)),
    if iteration == total:
        sys.stdout.write('\n')
    sys.stdout.flush()


if __name__ == "__main__":

    printProgress(1, 100, 'Progress:', 'Complete', 1, 50)
    printProgress(1, 100, 'Progress:', 'Complete', 1, 50)
    printProgress(1, 100, 'Progress:', 'Complete', 1, 50)
    printProgress(1, 100, 'Progress:', 'Complete', 1, 50)
    printProgress(1, 100, 'Progress:', 'Complete', 1, 50)
    printProgress(1, 100, 'Progress:', 'Complete', 1, 50)
    printProgress(1, 100, 'Progress:', 'Complete', 1, 50)

    # for i in range(0, 100):
    #     printProgress(i, 100, 'Progress:', 'Complete', 1, 50)
    #     os.system("pause")
