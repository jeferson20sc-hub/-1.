import { motion } from "framer-motion";
import { Sparkles } from "lucide-react";

import { Card, CardContent } from "@/components/ui/card";

interface Props {
  title: string;
  description: string;
}

export function ComingSoonView({ title, description }: Props) {
  return (
    <Card>
      <CardContent className="p-12 text-center">
        <motion.div
          animate={{ rotate: [0, 8, -8, 0] }}
          transition={{ duration: 4, repeat: Infinity }}
          className="size-16 rounded-2xl bg-gradient-to-br from-primary to-primary/60 grid place-items-center shadow-xl shadow-primary/30 mx-auto mb-5"
        >
          <Sparkles className="size-8 text-primary-foreground" strokeWidth={2.4} />
        </motion.div>
        <h2 className="text-xl font-bold tracking-tight">{title}</h2>
        <p className="text-sm text-muted-foreground mt-2 max-w-md mx-auto">
          {description}
        </p>
      </CardContent>
    </Card>
  );
}
