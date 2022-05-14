package readExcelFileGenerateData;

/**
 * 对 Excel 行数据的描述
 */
public class ExcelData {

    // 背景
    private String background;

    // 场景
    private String scenes;

    // 目标
    private String target;

    // 要点
    private String point;

    // 是否和当前习惯保持一致
    private boolean habit;

    // 行为
    private String behavior;

    // 价值
    private String value;

    public String getBackground() {
        return background;
    }

    public void setBackground(String background) {
        this.background = background;
    }

    public String getScenes() {
        return scenes;
    }

    public void setScenes(String scenes) {
        this.scenes = scenes;
    }

    public String getTarget() {
        return target;
    }

    public void setTarget(String target) {
        this.target = target;
    }

    public String getPoint() {
        return point;
    }

    public void setPoint(String point) {
        this.point = point;
    }

    public boolean isHabit() {
        return habit;
    }

    public void setHabit(boolean habit) {
        this.habit = habit;
    }

    public String getBehavior() {
        return behavior;
    }

    public void setBehavior(String behavior) {
        this.behavior = behavior;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }
}
