<?php
/**
 * Created by JetBrains PhpStorm.
 * User: bassimons
 * Date: 9/20/13
 * Time: 1:07 PM
 * To change this template use File | Settings | File Templates.
 */

namespace Bassim\BigXlsxBundle\Tests;


use Exception;
use PHPUnit_Framework_AssertionFailedError;
use PHPUnit_Framework_Test;
use PHPUnit_Framework_TestSuite;

class ExecutionTimeListener implements \PHPUnit_Framework_TestListener
{
	public function endTest(\PHPUnit_Framework_Test $test, $time)
	{
		printf("Test '%s' ended and took %s seconds.\n",
			$test->getName(),
			$test->time()
		);
	}

	/**
	 * An error occurred.
	 *
	 * @param  PHPUnit_Framework_Test $test
	 * @param  Exception $e
	 * @param  float $time
	 */
	public function addError(PHPUnit_Framework_Test $test, Exception $e, $time)
	{
		// TODO: Implement addError() method.
	}

	/**
	 * A failure occurred.
	 *
	 * @param  PHPUnit_Framework_Test $test
	 * @param  PHPUnit_Framework_AssertionFailedError $e
	 * @param  float $time
	 */
	public function addFailure(PHPUnit_Framework_Test $test, PHPUnit_Framework_AssertionFailedError $e, $time)
	{
		// TODO: Implement addFailure() method.
	}

	/**
	 * Incomplete test.
	 *
	 * @param  PHPUnit_Framework_Test $test
	 * @param  Exception $e
	 * @param  float $time
	 */
	public function addIncompleteTest(PHPUnit_Framework_Test $test, Exception $e, $time)
	{
		// TODO: Implement addIncompleteTest() method.
	}

	/**
	 * Skipped test.
	 *
	 * @param  PHPUnit_Framework_Test $test
	 * @param  Exception $e
	 * @param  float $time
	 * @since  Method available since Release 3.0.0
	 */
	public function addSkippedTest(PHPUnit_Framework_Test $test, Exception $e, $time)
	{
		// TODO: Implement addSkippedTest() method.
	}

	/**
	 * A test suite started.
	 *
	 * @param  PHPUnit_Framework_TestSuite $suite
	 * @since  Method available since Release 2.2.0
	 */
	public function startTestSuite(PHPUnit_Framework_TestSuite $suite)
	{
		// TODO: Implement startTestSuite() method.
	}

	/**
	 * A test suite ended.
	 *
	 * @param  PHPUnit_Framework_TestSuite $suite
	 * @since  Method available since Release 2.2.0
	 */
	public function endTestSuite(PHPUnit_Framework_TestSuite $suite)
	{
		// TODO: Implement endTestSuite() method.
	}

	/**
	 * A test started.
	 *
	 * @param  PHPUnit_Framework_Test $test
	 */
	public function startTest(PHPUnit_Framework_Test $test)
	{
		// TODO: Implement startTest() method.
	}
}

