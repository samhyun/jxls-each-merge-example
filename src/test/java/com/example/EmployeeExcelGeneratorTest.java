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

public class EmployeeExcelGeneratorTest {

    private EmployeeExcelGenerator generator;

    private List<Department> departments;

    @BeforeEach
    void setUp() {
        this.generator = new EmployeeExcelGenerator();
        this.departments = List.of(
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
    }

    @Test
    public void excel_generate_test() {
        Assertions.assertDoesNotThrow(() -> {
            generator.generate(departments, "output1");
        });
    }
}
