import java.time.DateTimeException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.temporal.WeekFields;
import java.util.Arrays;

/**
 * 週番号を日付に変換するサンプル
 * 
 * @author business
 *
 */
public class ConvertToLocalDate {

	public static void main(String[] args) {
		String[] aaa = {"2028", "2056"};
		LocalDate ld = null;
		int year = 2056;
		int week = 54;
		WeekFields wf = WeekFields.SUNDAY_START;
		try {
			if (week == 54 && Arrays.asList(aaa).contains(String.valueOf(year))) {
				ld = LocalDate.parse(year + "-12-31");
			} else {
				ld = LocalDate.now()
						.with(wf.weekBasedYear(), year)
						.with(wf.weekOfWeekBasedYear(), week)
						.with(wf.dayOfWeek(), DayOfWeek.MONDAY.getValue());
				
			}
			System.out.println("develop追加");
			System.out.println(ld);
		} catch (DateTimeException e) {
			System.out.println("日付変換エラー");
		}
	}

}
