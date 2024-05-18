import java.time.LocalTime;
import java.time.format.DateTimeFormatter;

public class BroadcastsTime implements Comparable<BroadcastsTime> {
    private LocalTime time;

    public BroadcastsTime(String timeStr) {
        this.time = LocalTime.parse(timeStr, DateTimeFormatter.ofPattern("HH:mm"));
    }

    public byte hour() {
        return (byte) time.getHour();
    }

    public byte minutes() {
        return (byte) time.getMinute();
    }

    public boolean after(BroadcastsTime t) {
        return this.time.isAfter(t.time);
    }

    public boolean before(BroadcastsTime t) {
        return this.time.isBefore(t.time);
    }

    public boolean between(BroadcastsTime t1, BroadcastsTime t2) {
        return (this.time.equals(t1.time) || this.time.isAfter(t1.time)) &&
                (this.time.equals(t2.time) || this.time.isBefore(t2.time));
    }

    @Override
    public int compareTo(BroadcastsTime o) {
        return this.time.compareTo(o.time);
    }

    @Override
    public String toString() {
        return time.format(DateTimeFormatter.ofPattern("HH:mm"));
    }
}
