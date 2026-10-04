"""
Auto Scaling collector module for Auto Scaling Group inventory.

Retrieves all Auto Scaling Groups and their configuration details
including capacity settings, instance counts, and launch templates.
"""

from typing import Any, Dict, List, Optional

from botocore.exceptions import ClientError, BotoCoreError

from .base_collector import BaseCollector


class AutoScalingCollector(BaseCollector):
    """Collects Auto Scaling Group configuration data.

    For each ASG:
    - Group name and ARN
    - Min / Max / Desired capacity
    - Current instance count
    - Launch template or launch configuration name
    - Health check type and grace period
    - Availability zones
    """

    def collect(self) -> List[Dict[str, Any]]:
        """Collect Auto Scaling Group inventory.

        Returns:
            List of ASG detail dicts.
        """
        self.logger.info(
            "Collecting Auto Scaling Group data for region %s…", self.region
        )

        asg_client = self._get_client("autoscaling")
        groups: List[dict] = []

        try:
            paginator = asg_client.get_paginator(
                "describe_auto_scaling_groups"
            )
            for page in paginator.paginate():
                groups.extend(page.get("AutoScalingGroups", []))
        except (ClientError, BotoCoreError) as exc:
            self.logger.error(
                "Failed to describe Auto Scaling Groups: %s", exc
            )
            return []

        if not groups:
            self.logger.info("No Auto Scaling Groups found in %s.", self.region)
            return []

        self.logger.info("Found %d Auto Scaling Groups.", len(groups))

        results: List[Dict[str, Any]] = []
        for group in groups:
            results.append(self._parse_group(group))

        self.logger.info(
            "Auto Scaling collection complete — %d groups.", len(results)
        )
        return results

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_group(group: dict) -> Dict[str, Any]:
        """Parse a raw ASG description into a summary dict."""
        launch_template = AutoScalingCollector._get_launch_template(group)
        launch_config = group.get("LaunchConfigurationName", None)

        instances = group.get("Instances", [])
        instance_ids = [i.get("InstanceId", "") for i in instances]

        return {
            "name": group.get("AutoScalingGroupName", ""),
            "arn": group.get("AutoScalingGroupARN", ""),
            "min_size": group.get("MinSize", 0),
            "max_size": group.get("MaxSize", 0),
            "desired_capacity": group.get("DesiredCapacity", 0),
            "instance_count": len(instances),
            "instance_ids": instance_ids,
            "launch_template": launch_template,
            "launch_configuration": launch_config,
            "health_check_type": group.get("HealthCheckType", ""),
            "health_check_grace_period": group.get(
                "HealthCheckGracePeriod", 0
            ),
            "availability_zones": group.get("AvailabilityZones", []),
            "status": group.get("Status", ""),
        }

    @staticmethod
    def _get_launch_template(group: dict) -> Optional[str]:
        """Extract the launch template name from an ASG.

        Checks both ``LaunchTemplate`` and ``MixedInstancesPolicy``.
        """
        # Direct launch template
        lt = group.get("LaunchTemplate")
        if lt:
            return lt.get("LaunchTemplateName") or lt.get(
                "LaunchTemplateId", None
            )

        # Mixed instances policy
        mip = group.get("MixedInstancesPolicy", {})
        lts = mip.get("LaunchTemplate", {})
        lt_spec = lts.get("LaunchTemplateSpecification", {})
        if lt_spec:
            return lt_spec.get("LaunchTemplateName") or lt_spec.get(
                "LaunchTemplateId", None
            )

        return None
