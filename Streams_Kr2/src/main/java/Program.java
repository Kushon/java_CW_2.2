public class Program {
    private String channel;
    private BroadcastsTime time;
    private String title;


    public Program(String channel, BroadcastsTime time, String title) {
        this.channel = channel;
        this.time = time;
        this.title = title;
    }

    public String getChannel() {
        return channel;
    }
    public BroadcastsTime getTime() {
        return time;
    }
    public String getTitle() {
        return title;
    }
    @Override
    public String toString() {        return "Канал: " + channel + ", Время: " + time + ", Название: " + title;
    }
}


