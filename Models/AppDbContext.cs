using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace MODSemanal.Models;

public partial class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options)
        : base(options)
    {
    }

    public virtual DbSet<ExcessHoursDistribution> ExcessHoursDistribution { get; set; }

    public virtual DbSet<HoursDistributionDetail> HoursDistributionDetail { get; set; }

    public virtual DbSet<WeeklyPlan> WeeklyPlan { get; set; }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<ExcessHoursDistribution>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK__ExcessHo__3214EC0752D3C053");

            entity.Property(e => e.MaterialType)
                .HasMaxLength(10)
                .IsUnicode(false);
            entity.Property(e => e.Mod).HasColumnName("MOD");

            entity.HasOne(d => d.WeeklyPlan).WithMany(p => p.ExcessHoursDistribution)
                .HasForeignKey(d => d.WeeklyPlanId)
                .HasConstraintName("FK__ExcessHou__Weekl__398D8EEE");
        });

        modelBuilder.Entity<HoursDistributionDetail>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK__HoursDis__3214EC0786DD8B91");

            entity.Property(e => e.DistributionType)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.HasOne(d => d.Distribution).WithMany(p => p.HoursDistributionDetail)
                .HasForeignKey(d => d.DistributionId)
                .HasConstraintName("FK__HoursDist__Distr__3C69FB99");
        });

        modelBuilder.Entity<WeeklyPlan>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK__WeeklyPl__3214EC07D4BD6737");

            entity.Property(e => e.ExcessHoursPerPerson).HasColumnType("decimal(10, 2)");
            entity.Property(e => e.MaterialType)
                .HasMaxLength(10)
                .IsUnicode(false);
            entity.Property(e => e.Mod).HasColumnName("MOD");
            entity.Property(e => e.ProductivityTarget).HasColumnType("decimal(10, 2)");
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}