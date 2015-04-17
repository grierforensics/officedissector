#!/usr/bin/env python

"""Performance tests to ensure plugins remain within CPU and memory requirements."""

import os
import sys
import subprocess


def main():
    if len(sys.argv) != 2:
        print "\nUsage: " + sys.argv[0] + " MASTIFF_SOURCE_DIR\n\n"
        sys.exit(1)

    mastiff_dir = sys.argv[1]
    test_dir = os.path.abspath(os.path.dirname(__file__))
    os.chdir(test_dir)

    corpus_path = ['govdocs/', 'fraunhoferlibrary/']
    files = []
    for dir_ in corpus_path:
        for f in os.listdir(dir_):
            file_ = os.path.join(dir_, f)
            if os.path.isfile(file_):
                files.append(file_)

    sum_results = [0.0, 0.0, 0.0]
    max_results = [0.0, 0.0, 0.0]
    num_results = 0
    failed = []
    for doc in files:
        results = run_time_with_mastiff(doc, mastiff_dir, test_dir)
        for i in range(len(results)):
            sum_results[i] += results[i]
            if results[i] > max_results[i]:
                max_results[i] = results[i]
        num_results += 1
        size = os.path.getsize(doc)/1024.0
        totaltime = results[0] + results[1]
        if ((size/1024.0) <= 10) and (totaltime > 30):
            passorfail = '[FAIL]'
            failed.append((doc, size, totaltime))
        else:
            # passorfail = '[PASS]'
            passorfail = '[TIME_OK]' # RAM requirements in process of clarification
        print '[%s: Size:%6.1fKB User:%.2f System:%.2f Maxresident:%.0fKB (%.0fx)]%s' % ((doc, size) + tuple(results) + (results[2]/size, passorfail))

    avg_results = [s/num_results for s in sum_results]
    print '\n\n[Average: User:%.2f System:%.2f Maxresident:%.0fKB]' % tuple(avg_results)
    print '[Maximum: User:%.2f System:%.2f Maxresident:%.0fKB]\n' % tuple(max_results)

    if failed:
        for fail in failed:
            print '[FAILED] Document: %s Size: %.1fKB Total time: %.2f' % fail
    else:
        print '[PASSED]\n'


def run_time_with_mastiff(doc, mastiff_dir, test_dir):
    os.chdir(mastiff_dir)
    try:
        p = subprocess.Popen(['/usr/bin/time',
                              'mas.py',
                              os.path.join(test_dir, doc),
                              '>/dev/null'],
                             stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except OSError as e:
        if e.errno == os.errno.ENOENT:
            print '\nmas.py cannot be found.'
            print 'Please ensure that MASTIFF is installed.'
            sys.exit(1)
        else:
            raise
    out, err = p.communicate()
    output = err.split('\n')
    if 'Could not read any configuration files' in output[0]:
        print 'Most likely your argument does not point to'
        print 'a valid MASTIFF source directory.'
        sys.exit(1)
    time_output = [string for i, string in enumerate(output) if 'maxresident)k' in string][0].split()
    user_time = [string for i, string in enumerate(time_output) if 'user' in string][0].split('user')[0]
    system_time = [string for i, string in enumerate(time_output) if 'system' in string][0].split('system')[0]
    max_resident = [string for i, string in enumerate(time_output) if 'maxresident' in string][0].split('maxresident')[0]

    os.chdir(test_dir)

    return float(user_time), float(system_time), float(max_resident)

if __name__ == '__main__':
    main()

