public class Job {
    private String name;
    private String company;
    private String location;

    public Job(String name, String company, String location) {
        this.name = name;
        this.company = company;
        this.location = location;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getCompany() {
        return company;
    }

    public void setCompany(String company) {
        this.company = company;
    }

    public String getLocation() {
        return location;
    }

    public void setLocation(String location) {
        this.location = location;
    }

    @Override

    public String toString() {
        return
                "NAME   =   " + name + "  || " +
                " COMPANY   =   " + company + "  || " +
                " LOCATION   =   " + location
                ;
    }


}
