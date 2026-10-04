"""
WAF collector module for AWS WAFv2 Web ACL metrics.

Retrieves all Web ACLs (REGIONAL and CLOUDFRONT scopes) and their
CloudWatch request metrics: allowed, blocked, counted, CAPTCHA,
and challenge requests.
"""

from typing import Any, Dict, List

from botocore.exceptions import ClientError, BotoCoreError

from .base_collector import BaseCollector


class WAFCollector(BaseCollector):
    """Collects AWS WAFv2 Web ACL data and CloudWatch metrics.

    Scans both REGIONAL (current region) and CLOUDFRONT (us-east-1)
    scopes.  For each Web ACL, collects:
    - AllowedRequests
    - BlockedRequests
    - CountedRequests
    - CaptchaRequests (gracefully handles absence)
    - ChallengedRequests (gracefully handles absence)
    - Total requests (sum of allowed + blocked + counted)
    """

    def collect(self) -> List[Dict[str, Any]]:
        """Collect WAF Web ACL inventory and request metrics.

        Returns:
            List of Web ACL detail dicts.
        """
        self.logger.info("Collecting WAFv2 data…")

        results: List[Dict[str, Any]] = []

        # REGIONAL scope (current region)
        regional_acls = self._list_web_acls(scope="REGIONAL", region=self.region)
        for acl in regional_acls:
            metrics = self._collect_acl_metrics(
                acl_name=acl["Name"],
                region=self.region,
                scope="REGIONAL",
            )
            metrics["scope"] = "REGIONAL"
            metrics["acl_id"] = acl.get("Id", "")
            results.append(metrics)

        # CLOUDFRONT scope (must use us-east-1)
        cloudfront_acls = self._list_web_acls(scope="CLOUDFRONT", region="us-east-1")
        for acl in cloudfront_acls:
            metrics = self._collect_acl_metrics(
                acl_name=acl["Name"],
                region="us-east-1",
                scope="CLOUDFRONT",
            )
            metrics["scope"] = "CLOUDFRONT"
            metrics["acl_id"] = acl.get("Id", "")
            results.append(metrics)

        self.logger.info(
            "WAF collection complete — %d Web ACLs (%d regional, %d CloudFront).",
            len(results),
            len(regional_acls),
            len(cloudfront_acls),
        )
        return results

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _list_web_acls(self, scope: str, region: str) -> List[dict]:
        """List all Web ACLs for a given scope and region."""
        wafv2 = self._get_client("wafv2", region_override=region)
        acls: List[dict] = []
        try:
            # list_web_acls supports NextMarker pagination
            params: Dict[str, Any] = {"Scope": scope, "Limit": 100}
            while True:
                response = wafv2.list_web_acls(**params)
                acls.extend(response.get("WebACLs", []))
                next_marker = response.get("NextMarker")
                if not next_marker:
                    break
                params["NextMarker"] = next_marker
        except (ClientError, BotoCoreError) as exc:
            self.logger.warning(
                "Failed to list WAF Web ACLs (scope=%s, region=%s): %s",
                scope,
                region,
                exc,
            )
        return acls

    def _collect_acl_metrics(
        self, acl_name: str, region: str, scope: str
    ) -> Dict[str, Any]:
        """Collect CloudWatch request metrics for a single Web ACL."""
        data: Dict[str, Any] = {"name": acl_name}

        # Dimension set used for WAFv2 metrics
        metric_region = region if scope == "REGIONAL" else "us-east-1"
        dims = [
            {"Name": "WebACL", "Value": acl_name},
            {"Name": "Region", "Value": metric_region},
            {"Name": "Rule", "Value": "ALL"},
        ]

        # Standard request metrics
        standard_metrics = [
            ("AllowedRequests", "allowed_requests"),
            ("BlockedRequests", "blocked_requests"),
            ("CountedRequests", "counted_requests"),
        ]

        for metric_name, key in standard_metrics:
            datapoints = self._get_cloudwatch_metric(
                namespace="AWS/WAFV2",
                metric_name=metric_name,
                dimensions=dims,
                statistics=["Sum"],
                period=300, # 5 minutes
            )
            if datapoints:
                total = sum(dp.get("Sum", 0.0) for dp in datapoints if dp.get("Sum") is not None)
                data[key] = int(round(total))
            else:
                data[key] = 0

        # CAPTCHA and Challenge — may not exist
        optional_metrics = [
            ("CaptchaRequests", "captcha_requests"),
            ("ChallengeRequests", "challenge_requests"),
        ]

        for metric_name, key in optional_metrics:
            datapoints = self._get_cloudwatch_metric(
                namespace="AWS/WAFV2",
                metric_name=metric_name,
                dimensions=dims,
                statistics=["Sum"],
                period=300, # 5 minutes
            )
            if datapoints:
                total = sum(dp.get("Sum", 0.0) for dp in datapoints if dp.get("Sum") is not None)
                data[key] = int(round(total))
            else:
                data[key] = 0

        # Derived total
        data["total_requests"] = (
            data["allowed_requests"]
            + data["blocked_requests"]
            + data["counted_requests"]
            + data["captcha_requests"]
            + data["challenge_requests"]
        )

        return data
