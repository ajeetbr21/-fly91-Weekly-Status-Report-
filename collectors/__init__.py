"""
AWS Weekly BAU Report — Collectors Package.

This package provides data collection classes for each AWS service
covered by the weekly BAU report.  All collectors inherit from
:class:`BaseCollector` and implement a ``collect()`` method that
returns a service-specific dictionary of results.

Usage::

    from collectors import (
        CostCollector,
        EC2Collector,
        ELBCollector,
        WAFCollector,
        S3Collector,
        RDSCollector,
        AutoScalingCollector,
    )
"""

from .base_collector import BaseCollector
from .cost_collector import CostCollector
from .ec2_collector import EC2Collector
from .elb_collector import ELBCollector
from .waf_collector import WAFCollector
from .s3_collector import S3Collector
from .rds_collector import RDSCollector
from .autoscaling_collector import AutoScalingCollector
from .inspector_collector import InspectorCollector
from .guardduty_collector import GuardDutyCollector

__all__ = [
    "BaseCollector",
    "CostCollector",
    "EC2Collector",
    "ELBCollector",
    "WAFCollector",
    "S3Collector",
    "RDSCollector",
    "AutoScalingCollector",
    "InspectorCollector",
    "GuardDutyCollector",
]
