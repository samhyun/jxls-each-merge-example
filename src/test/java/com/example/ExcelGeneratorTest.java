package com.example;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.Month;
import java.time.ZoneId;
import java.util.Date;
import java.util.List;

public class ExcelGeneratorTest {

    private ExcelGenerator generator;

    private List<Department> departments;

    private List<Company> companies;

    private List<Country> countries;
    @BeforeEach
    void setUp() {
        departments = List.of(
                Department.builder()
                        .name("개발 1 팀")
                        .employees(List.of(
                                Employee.builder()
                                        .name("홍길동")
                                        .birthDate(Date.from(LocalDate.of(1978, Month.JANUARY, 1).atStartOfDay(ZoneId.systemDefault()).toInstant()))
                                        .bonus(BigDecimal.valueOf(10000))
                                        .payment(BigDecimal.valueOf(200000))
                                        .build(),
                                Employee.builder()
                                        .name("이순신")
                                        .birthDate(Date.from(LocalDate.of(1989, Month.DECEMBER, 22).atStartOfDay(ZoneId.systemDefault()).toInstant()))
                                        .bonus(BigDecimal.valueOf(120000))
                                        .payment(BigDecimal.valueOf(2200000))
                                        .build()
                        ))
                        .build(),
                Department.builder()
                        .name("개발 2 팀")
                        .employees(List.of(
                                Employee.builder()
                                        .name("차범근")
                                        .birthDate(Date.from(LocalDate.of(1955, Month.AUGUST, 4).atStartOfDay(ZoneId.systemDefault()).toInstant()))
                                        .bonus(BigDecimal.valueOf(410000))
                                        .payment(BigDecimal.valueOf(2000000))
                                        .build(),
                                Employee.builder()
                                        .name("박지성")
                                        .birthDate(Date.from(LocalDate.of(1981, Month.JUNE, 30).atStartOfDay(ZoneId.systemDefault()).toInstant()))
                                        .bonus(BigDecimal.valueOf(320000))
                                        .payment(BigDecimal.valueOf(1500000))
                                        .build(),
                                Employee.builder()
                                        .name("손흥민")
                                        .birthDate(Date.from(LocalDate.of(1992, Month.MARCH, 11).atStartOfDay(ZoneId.systemDefault()).toInstant()))
                                        .bonus(BigDecimal.valueOf(120000))
                                        .payment(BigDecimal.valueOf(1000000))
                                        .build()
                        ))
                        .build()
        );

        companies = List.of(
            Company.builder()
                    .name("좋은 회사")
                    .departments(departments)
                    .build(),
            Company.builder()
                    .name("나쁜 회사")
                    .departments(departments)
                    .build(),
                Company.builder()
                        .name("이상한 회사")
                        .departments(departments)
                        .build()
        );

        countries = List.of(
                Country.builder()
                        .name("대한민국")
                        .companies(companies)
                        .build(),
                Country.builder()
                        .name("미국")
                        .companies(companies)
                        .build(),
                Country.builder()
                        .name("독일")
                        .companies(companies)
                        .build()
        );
    }
    @Test
    public void employees_excel_generate_test() {
        this.generator = new ExcelGenerator("object_collection_template.xls");
        this.generator.addMappingValue("employees", departments.get(0).getEmployees());

        Assertions.assertDoesNotThrow(() -> {
            generator.generate("output-employees.xlsx");
        });
    }

    @Test
    public void department_excel_generate_test() {
        this.generator = new ExcelGenerator("each-merge-template.xlsx");
        this.generator.addMappingValue("departments", departments);

        Assertions.assertDoesNotThrow(() -> {
            generator.generate("output.xlsx");
        });
    }

    @Test
    public void nested_each_merge_generate_test() {
        this.generator = new ExcelGenerator("nested-each-merge-template.xlsx");
        this.generator.addMappingValue("companies", companies);

        Assertions.assertDoesNotThrow(() -> {
            generator.generate("nested-output.xlsx");
        });
    }

    @Test
    public void nested_nested_each_merge_generate_test() {
        this.generator = new ExcelGenerator("nested-nested-each-merge-template.xlsx");
        this.generator.addMappingValue("countries", countries);

        Assertions.assertDoesNotThrow(() -> {
            generator.generate("nested-nested-output.xlsx");
        });
    }

    @Test
    public void abnormal_nested_nested_each_merge_generate_test() {
        this.generator = new ExcelGenerator("abnormal-nested-nested-each-merge-template.xlsx");
        this.generator.addMappingValue("countries", countries);

        Assertions.assertDoesNotThrow(() -> {
            generator.generate("abnormal-nested-nested-output.xlsx");
        });
    }

}
