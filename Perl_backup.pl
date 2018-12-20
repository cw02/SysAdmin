#!/usr/bin/perl
#################################################################################
# AUTHOR:                   Christopher Weller
# CONTACT E-MAIL ADDRESS:   christopher.e.weller@gmail.com
# INITIAL DATE:             August 4, 2011
# VERSION:                  1.1print "Backup failed at: $theTime\n";
#                           print "Error is: $@\n";
# PURPOSE:                  Recursive copy of one directory to another, 
#                           for backup purposes
# COMMENTS:                 None
################################################################################
#                           Revision History
#
#   INITIALS     DATE       REASON
#______________________________________________________________________________
#    CW        08/04/2011   In-service date
#    CW        08/06/2011   Added logging of the backup process
#
#
#
################################################################################
#
#                              DISCLAIMER
#
# THIS SOFTWARE IS PROVIDED "AS IS" AND ANY EXPRESSED OR IMPLIED WARRANTIES,
# INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
# FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE REGENTS
# OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
# EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT
# OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.
#
#use strict;
#use 5.10.0;
use warnings;
use Archive::Zip qw/ :ERROR_CODES :CONSTANTS/;

                 # Define the locations to copy from and to.
$from = "/Users/Christopher";                            # location to be backed up
$to = "/Volumes/Iomega_HDD/Backup";          # location to place the backup files
$obj = Archive::Zip->new();
@months = qw(Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec);
@weekDays = qw(Sun Mon Tue Wed Thu Fri Sat Sun);
open (LOGFILE, '>>Backup.log');


                # Main execution portion of the backup script

eval {
                # Run the backup operation
	print "Backup is starting\n";
	($second, $minute, $hour, $dayOfMonth, $month, $yearOffset, $dayOfWeek, $dayOfYear, $daylightSavings) = localtime();
	$year = 1900 + $yearOffset;
	$theTime = "$hour:$minute:$second, $weekDays[$dayOfWeek] $months[$month] $dayOfMonth, $year";
	print LOGFILE "Backup started at: $theTime\n";
	print LOGFILE "Content-type: text/htmlnn\n";
	print "Backup started at: $theTime\n";
	print "Content-type: text/htmlnn\n";
	`cp -Rf $from $to`;
	($second, $minute, $hour, $dayOfMonth, $month, $yearOffset, $dayOfWeek, $dayOfYear, $daylightSavings) = localtime();
	$year = 1900 + $yearOffset;
	$theTime = "$hour:$minute:$second, $weekDays[$dayOfWeek] $months[$month] $dayOfMonth, $year";
	print "Backup finished at: $theTime\n";
	print LOGFILE "Backup finished at: $theTime\n";
	close (LOGFILE);
};


                   #  This is the error handling portion of the backup script.
                   #  Hopefully the this thread never runs.

if ($@){
	($second, $minute, $hour, $dayOfMonth, $month, $yearOffset, $dayOfWeek, $dayOfYear, $daylightSavings) = localtime();
	$year = 1900 + $yearOffset;
	$theTime = "$hour:$minute:$second, $weekDays[$dayOfWeek] $months[$month] $dayOfMonth, $year";
	print LOGFILE "Backup failed at: $theTime\n";
	print LOGFILE "Error is: $@\n";
	print "Backup failed at: $theTime\n";
	print "Error is: $@\n";
	close (LOGFILE);
	
};