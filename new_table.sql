DROP TABLE IF EXISTS `devices`;
 CREATE TABLE `devices`  (
   `id` int NOT NULL AUTO_INCREMENT,
   `material_id` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NOT NULL,
   `spare_count` int NULL DEFAULT 0,
   `unit` varchar(50) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL DEFAULT NULL,
   `remark` text CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL,
   `description` text CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL COMMENT '物料描述',
   `status` enum('白名单','黑名单') CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NOT NULL,
   `created_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP,
   `updated_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
   PRIMARY KEY (`id`) USING BTREE,
   UNIQUE INDEX `unique_material_id`(`material_id` ASC) USING BTREE,
   INDEX `idx_material_id_status`(`material_id` ASC, `status` ASC) USING BTREE
 ) ENGINE = InnoDB AUTO_INCREMENT = 7075 CHARACTER SET = utf8mb4 COLLATE = utf8mb4_general_ci ROW_FORMAT = Dynamic;

 SET FOREIGN_KEY_CHECKS = 1;